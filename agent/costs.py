"""Where the money went.

Until now a run reported ONE number -- cost-weighted tokens -- and that
number exists to feed the spending cap, not to explain anything. It cannot
answer the question you actually ask after a $65 surprise: *what* spent it.
OCR or testing? Which test step? Which turn? Fresh input, or the same
context being re-read every turn because caching silently broke?

So this records every billed API call with its four token kinds kept
separate, tagged with what it was for, and converts to dollars at the
published per-model rates. Nothing here participates in the spending cap
(agent/loop.py's own running total still does that, unchanged) -- this is
pure observation, deliberately kept off the control path so a bug in cost
REPORTING can never abort a run.

Two things it can show that a single weighted total cannot:

1. Cache health. cache_read at ~1/10th price is the whole reason a
   multi-turn step is affordable. If cache reads are near zero across turns
   2..n, something is invalidating the prefix and the run costs ~10x what
   it should -- visible here as a fresh-input column that keeps growing
   instead of a cache-read one.

2. Prompt composition. Turn 1 is the expensive turn, and it is expensive
   because of what we put in it: PY excerpts, the CY evidence inventory,
   the sample roster, the system prompt, the tool schemas. Section sizes
   are apportioned against the MEASURED turn-1 input rather than estimated
   at N characters per token, so the parts always sum to the real billed
   number -- see prompt_mix_rows().
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

# Published list prices, $ per million tokens: (fresh input, output).
# Matched by substring against the model id, most specific first, because
# the deployment name is whatever the Foundry resource calls it.
#
# Microsoft Foundry bills Claude at standard Anthropic API rates, so these
# are the right numbers for this app's deployment.
_PRICES: list[tuple[str, tuple[float, float]]] = [
    ("fable", (10.00, 50.00)),
    ("mythos", (10.00, 50.00)),
    ("opus", (5.00, 25.00)),
    ("sonnet-4-6", (3.00, 15.00)),
    ("sonnet", (2.00, 10.00)),
    ("haiku", (1.00, 5.00)),
]
_FALLBACK = ("opus", (5.00, 25.00))

# Relative to fresh input. These ratios are stable across the model line,
# which is why the spending cap in agent/loop.py can use them as pure
# weights without knowing the model. (Claude Fable 5.1 reads cache cheaper
# than 0.1x; using 0.1x there over-states the bill, which is the safe
# direction for a cost display to be wrong in.)
CACHE_WRITE_MULTIPLIER = 1.25
CACHE_READ_MULTIPLIER = 0.1
OUTPUT_MULTIPLIER_NOTE = "output is priced 5x input on every current model"


def prices_for(model: str) -> tuple[str, float, float, float, float]:
    """(matched family, fresh input, cache write, cache read, output) in
    $ per million tokens. An unrecognised deployment name falls back to
    Opus rates -- the family this app actually runs on, and the one that
    over- rather than under-states a Sonnet/Haiku bill.
    """
    lowered = (model or "").lower()
    family, (fresh, output) = _FALLBACK
    for key, price in _PRICES:
        if key in lowered:
            family, (fresh, output) = key, price
            break
    return family, fresh, fresh * CACHE_WRITE_MULTIPLIER, fresh * CACHE_READ_MULTIPLIER, output


def _field(usage: Any, name: str) -> int:
    """One usage field, coerced to an int.

    The ``or 0`` is not decoration. On the real SDK the cache fields are
    Optional[int]: the attribute EXISTS but is None when unpopulated, so
    getattr's default never fires and a bare sum raises TypeError on the
    first live turn. That exact bug reached a real run once already.
    """
    return int(getattr(usage, name, 0) or 0)


@dataclass(frozen=True)
class UsageRecord:
    """One billed API call, with the four token kinds kept apart."""

    group: str  # "OCR", or a test_step_id -- the coarse "where"
    label: str  # the fine "where": which page, or which tools that turn called
    turn: int  # 0 for OCR pages, which have no turn of their own
    input_tokens: int
    cache_write_tokens: int
    cache_read_tokens: int
    output_tokens: int

    @property
    def raw_tokens(self) -> int:
        """What a provider dashboard shows: every token at face value."""
        return self.input_tokens + self.cache_write_tokens + self.cache_read_tokens + self.output_tokens

    def dollars(self, model: str) -> float:
        _, fresh, cache_write, cache_read, output = prices_for(model)
        return (
            self.input_tokens * fresh
            + self.cache_write_tokens * cache_write
            + self.cache_read_tokens * cache_read
            + self.output_tokens * output
        ) / 1_000_000


@dataclass
class GroupTotal:
    group: str
    calls: int
    input_tokens: int
    cache_write_tokens: int
    cache_read_tokens: int
    output_tokens: int
    dollars: float

    @property
    def raw_tokens(self) -> int:
        return self.input_tokens + self.cache_write_tokens + self.cache_read_tokens + self.output_tokens


class TokenLedger:
    """Every billed call of one control run, in the order they happened.

    Deliberately append-only and side-effect free: run_control passes one
    of these down through OCR and every test step, and callers read it
    after the fact. It never raises on a malformed usage object -- a run
    that finished must not fail while reporting what it cost.
    """

    def __init__(self, model: str) -> None:
        self.model = model
        self.records: list[UsageRecord] = []
        # {group: {section name: characters}} -- see note_prompt_mix.
        self.prompt_mix: dict[str, dict[str, int]] = {}

    # -- recording ---------------------------------------------------------

    def record(self, usage: Any, *, group: str, label: str, turn: int = 0) -> UsageRecord:
        rec = UsageRecord(
            group=group,
            label=label,
            turn=turn,
            input_tokens=_field(usage, "input_tokens"),
            cache_write_tokens=_field(usage, "cache_creation_input_tokens"),
            cache_read_tokens=_field(usage, "cache_read_input_tokens"),
            output_tokens=_field(usage, "output_tokens"),
        )
        self.records.append(rec)
        return rec

    def note_prompt_mix(self, group: str, sections: dict[str, int]) -> None:
        """Character counts of the named pieces of a step's first request.

        Recorded once per step, before the call. Characters (not tokens)
        because that is what we can measure locally and for free; they get
        converted into real token numbers in prompt_mix_rows() by
        apportioning the MEASURED turn-1 input across them, so the reported
        parts always add up to what was actually billed.
        """
        self.prompt_mix[group] = dict(sections)

    # -- totals ------------------------------------------------------------

    @property
    def dollars(self) -> float:
        return sum(r.dollars(self.model) for r in self.records)

    @property
    def raw_tokens(self) -> int:
        return sum(r.raw_tokens for r in self.records)

    def kind_rows(self) -> list[tuple[str, int, float, float]]:
        """(token kind, tokens, dollars, % of spend), biggest first.

        This is the cache-health view. On a healthy multi-turn step, cache
        reads are the largest token count and one of the smallest dollar
        amounts -- if those two facts don't both hold, caching isn't
        working and the run is paying full price to re-read its own
        context every turn.
        """
        _, fresh, cache_write, cache_read, output = prices_for(self.model)
        kinds = [
            ("Fresh input (never cached)", sum(r.input_tokens for r in self.records), fresh),
            ("Cache writes", sum(r.cache_write_tokens for r in self.records), cache_write),
            ("Cache reads", sum(r.cache_read_tokens for r in self.records), cache_read),
            ("Output", sum(r.output_tokens for r in self.records), output),
        ]
        total = self.dollars
        rows = [(name, tok, tok * rate / 1_000_000) for name, tok, rate in kinds]
        return [
            (name, tok, usd, (100.0 * usd / total) if total else 0.0)
            for name, tok, usd in sorted(rows, key=lambda r: -r[2])
        ]

    def group_rows(self) -> list[GroupTotal]:
        """One row per phase -- OCR, then each test step -- in run order.

        The first question after a surprising bill is always "which part of
        it", and this is that answer: OCR happens outside the per-step cap
        entirely, and a single pathological step can dominate a control.
        """
        order: list[str] = []
        acc: dict[str, GroupTotal] = {}
        for r in self.records:
            if r.group not in acc:
                order.append(r.group)
                acc[r.group] = GroupTotal(r.group, 0, 0, 0, 0, 0, 0.0)
            g = acc[r.group]
            g.calls += 1
            g.input_tokens += r.input_tokens
            g.cache_write_tokens += r.cache_write_tokens
            g.cache_read_tokens += r.cache_read_tokens
            g.output_tokens += r.output_tokens
            g.dollars += r.dollars(self.model)
        return [acc[name] for name in order]

    def call_rows(self, group: str | None = None) -> list[tuple[UsageRecord, float]]:
        """(record, dollars) for each call, optionally one group only --
        the turn-by-turn view, where a single runaway turn shows up.
        """
        return [
            (r, r.dollars(self.model))
            for r in self.records
            if group is None or r.group == group
        ]

    def prompt_mix_rows(self, group: str) -> list[tuple[str, int, int, float]]:
        """(section, characters, approx tokens, approx dollars) for what
        filled a step's first request, largest first.

        Apportioned across the measured turn-1 input (fresh + cache write,
        which together are everything sent that turn), so the numbers sum
        to the real billed figure rather than to a chars/4 guess. The split
        BETWEEN sections is still proportional-by-characters and therefore
        approximate -- dense tabular text tokenizes worse than prose -- but
        it is more than accurate enough to answer "is it the PY excerpts?",
        which is the only question anyone asks of it.
        """
        sections = self.prompt_mix.get(group)
        if not sections:
            return []
        total_chars = sum(sections.values())
        if total_chars <= 0:
            return []

        first = next((r for r in self.records if r.group == group), None)
        billed = (first.input_tokens + first.cache_write_tokens) if first else 0

        _, fresh, cache_write, _, _ = prices_for(self.model)
        # Blend the two rates by how the turn actually split, so a cached
        # section isn't priced as if it were fresh.
        if first and billed:
            rate = (first.input_tokens * fresh + first.cache_write_tokens * cache_write) / billed
        else:
            rate = fresh

        rows = []
        for name, chars in sections.items():
            share = chars / total_chars
            tokens = round(billed * share)
            rows.append((name, chars, tokens, tokens * rate / 1_000_000))
        return sorted(rows, key=lambda r: -r[1])

    # -- text report (CLI) -------------------------------------------------

    def summary_lines(self) -> list[str]:
        family, fresh, _, _, output = prices_for(self.model)
        lines = [
            f"TOKEN BURN -- ${self.dollars:,.2f} across {len(self.records)} API call(s), "
            f"{self.raw_tokens:,} raw tokens",
            f"  priced as {family} (${fresh:.2f}/${output:.2f} per Mtok in/out); "
            f"cache writes {CACHE_WRITE_MULTIPLIER}x input, cache reads {CACHE_READ_MULTIPLIER}x",
            "",
            "  Where it went:",
        ]
        total = self.dollars
        for g in self.group_rows():
            pct = (100.0 * g.dollars / total) if total else 0.0
            lines.append(
                f"    {g.group:<24} ${g.dollars:>8,.2f}  {pct:>5.1f}%  "
                f"({g.calls} call(s), {g.raw_tokens:,} raw tokens)"
            )
        lines += ["", "  By token kind:"]
        for name, tok, usd, pct in self.kind_rows():
            lines.append(f"    {name:<28} {tok:>10,} tok  ${usd:>8,.2f}  {pct:>5.1f}%")
        return lines

    def report_lines(self) -> list[str]:
        """The full record: the summary, plus every call and what filled
        each step's first prompt. This is what gets written next to the
        results, so a control that looked expensive can be explained later
        instead of re-run to find out why.
        """
        lines = list(self.summary_lines())
        for g in self.group_rows():
            lines += ["", f"  {g.group} -- ${g.dollars:,.4f} across {g.calls} call(s):"]
            lines.append(
                f"    {'turn':<5} {'what it did':<40} {'fresh':>8} {'cwrite':>8} "
                f"{'cread':>8} {'output':>8}  {'cost':>9}"
            )
            for rec, usd in self.call_rows(g.group):
                lines.append(
                    f"    {rec.turn or '-':<5} {rec.label[:40]:<40} {rec.input_tokens:>8,} "
                    f"{rec.cache_write_tokens:>8,} {rec.cache_read_tokens:>8,} "
                    f"{rec.output_tokens:>8,}  ${usd:>8,.4f}"
                )
            mix = self.prompt_mix_rows(g.group)
            if mix:
                lines.append(f"    what filled {g.group}'s first prompt:")
                for name, chars, tokens, usd in mix:
                    lines.append(
                        f"      {name:<32} {chars:>8,} chars  ~{tokens:>7,} tok  ${usd:>8,.4f}"
                    )
        return lines
