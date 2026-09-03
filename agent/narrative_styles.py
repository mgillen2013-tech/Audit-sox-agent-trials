"""Two ways to write the narrative box, so they can be compared.

This is deliberately an EXPERIMENT, not a decision. Both styles are
implemented, both are reachable by a flag, and neither is privileged in
the code -- the point is to run real controls through each and find out
which a reviewer prefers, or whether the answer is some third thing.

MODEL   - what the agent wrote, capped to a couple of sentences. Every
          sample's box reads differently, because the model rephrases per
          item. Fluent, and each box says something specific.

TEMPLATE - the same skeleton for every sample, with only the sample's own
          values substituted. Box to box, the sentences are identical and
          only the invoice number, amount, date and names move.

The case for TEMPLATE, which is why it is worth testing at all: a
reviewer reads five sample tabs in a row. When the prose is identical,
the only thing their eye has to do is check the values -- differences
jump out because everything else is constant. When each box is phrased
differently they have to read all five properly to be sure nothing was
said differently for a reason. Prior-year workpapers are written this way
by human testers, who arrived at it for the same reason.

The cost of TEMPLATE is that a genuinely unusual sample -- a credit memo,
a partial payment, an exception with a story -- has nowhere to say so.
That is what the summary narrative is for, and why this module moves the
model's prose there rather than deleting it.
"""

from __future__ import annotations

from typing import Literal

from agent.ooxml.models import NarrativeParagraph, NarrativeRun

NarrativeStyle = Literal["model", "template"]

# The fixed skeleton. Every sample's box opens and closes with these exact
# words -- no substitution, no rephrasing -- so anything a reviewer sees
# differing between two tabs is a real difference in the evidence.
_OPENING = "IA agreed the sampled item to supporting documentation and verified the following:"
_CLOSING_SATISFIED = "Test step satisfied. No exceptions noted."
_CLOSING_EXCEPTION = "Exception noted -- see the Summary tab."


def build_narrative(
    style: NarrativeStyle,
    *,
    step_label: str,
    model_narrative: str,
    legend: list[tuple[str, str, str]],
    satisfied: bool,
    lead_sentences,
) -> list[NarrativeParagraph]:
    """The narrative box for one sampled item.

    legend: (letter, attribute, value_observed) per tickmark, in order.
    lead_sentences: the truncator applied to model prose -- passed in so
    this module does not import back into the bridge.
    """
    paras: list[NarrativeParagraph] = []

    if step_label:
        paras.append(
            NarrativeParagraph(runs=[NarrativeRun(text=step_label, bold=True, color="BLACK")])
        )

    if style == "template":
        paras.append(NarrativeParagraph(runs=[NarrativeRun(text=_OPENING, color="BLACK")]))
    else:
        lead = lead_sentences(model_narrative)
        if lead:
            paras.append(NarrativeParagraph(runs=[NarrativeRun(text=lead, color="BLACK")]))

    paras.append(NarrativeParagraph())

    for letter, attribute, value in legend:
        # The legend is identical in shape under BOTH styles. It is already
        # the most template-like part of the box, and the thing a reviewer
        # actually ties to the red marks on the exhibits.
        paras.append(
            NarrativeParagraph(
                runs=[
                    NarrativeRun(text=f"{letter} - ", bold=True, color="RED"),
                    NarrativeRun(text=f"{attribute}: {value}", color="BLACK"),
                ]
            )
        )

    if style == "template":
        paras.append(NarrativeParagraph())
        paras.append(
            NarrativeParagraph(
                runs=[
                    NarrativeRun(
                        text=_CLOSING_SATISFIED if satisfied else _CLOSING_EXCEPTION,
                        bold=True,
                        color="BLACK" if satisfied else "RED",
                    )
                ]
            )
        )

    return paras
