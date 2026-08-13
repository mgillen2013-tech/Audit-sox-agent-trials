"""Shared fake-Anthropic-client helpers for tests that drive agent/loop.py
without a network call or an API key. Used by test_loop.py and
test_run_control.py.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any


def tool_use(id_: str, name: str, input_: dict) -> SimpleNamespace:
    return SimpleNamespace(type="tool_use", id=id_, name=name, input=input_)


def text_block(t: str) -> SimpleNamespace:
    return SimpleNamespace(type="text", text=t)


def fake_response(content: list, stop_reason: str = "tool_use") -> SimpleNamespace:
    return SimpleNamespace(content=content, stop_reason=stop_reason)


@dataclass
class FakeClient:
    """Matches agent.loop.AnthropicClientLike: a bare create_message(**kwargs)."""

    responses: list[Any]
    calls: list[dict] = field(default_factory=list)
    _index: int = 0

    def create_message(self, **kwargs) -> Any:
        self.calls.append(kwargs)
        r = self.responses[self._index]
        self._index += 1
        return r
