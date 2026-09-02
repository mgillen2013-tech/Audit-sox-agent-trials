"""Shared fake-Anthropic-client helpers for tests that drive agent/loop.py
without a network call or an API key. Used by test_loop.py and
test_run_control.py.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any


def tool_use(id_: str, name: str, input_: dict) -> SimpleNamespace:
    return SimpleNamespace(type="tool_use", id=id_, name=name, input=input_)


def text_block(t: str) -> SimpleNamespace:
    return SimpleNamespace(type="text", text=t)


def fake_usage(
    input_tokens: int = 0,
    output_tokens: int = 0,
    cache_creation_input_tokens: "int | None" = None,
    cache_read_input_tokens: "int | None" = None,
) -> SimpleNamespace:
    # The cache fields default to None, NOT 0, deliberately: on the real SDK
    # they are Optional[int] and come back as None when unpopulated. An
    # all-int fake masked exactly that -- _usage_tokens shipped with a
    # `getattr(usage, field, 0)` that would have raised TypeError on the
    # first real turn, and every test passed. The double must be at least as
    # hostile as the real SDK.
    return SimpleNamespace(
        input_tokens=input_tokens,
        output_tokens=output_tokens,
        cache_creation_input_tokens=cache_creation_input_tokens,
        cache_read_input_tokens=cache_read_input_tokens,
    )


def fake_response(content: list, stop_reason: str = "tool_use", usage: SimpleNamespace | None = None) -> SimpleNamespace:
    # usage defaults to all-zero -- existing tests that don't care about the
    # token-budget circuit breaker (agent.loop.MAX_TOTAL_TOKENS) never
    # accidentally trip it just by running a few turns.
    return SimpleNamespace(content=content, stop_reason=stop_reason, usage=usage or fake_usage())


@dataclass
class FakeClient:
    """Matches agent.loop.AnthropicClientLike: a bare create_message(**kwargs)."""

    responses: list[Any]
    calls: list[dict] = field(default_factory=list)
    _index: int = 0

    def create_message(self, **kwargs) -> Any:
        # Deep-copy messages[] -- agent/loop.py mutates the SAME list AND
        # the same nested block dicts across turns (appending messages, and
        # both adding and later stripping cache_control markers in place as
        # the conversation grows), so storing kwargs as-is -- or even a
        # shallow list-copy -- would make every entry in self.calls
        # silently reflect the FINAL state instead of what was actually
        # sent at call time.
        snapshot = {**kwargs, "messages": copy.deepcopy(kwargs["messages"])}
        self.calls.append(snapshot)
        r = self.responses[self._index]
        self._index += 1
        return r
