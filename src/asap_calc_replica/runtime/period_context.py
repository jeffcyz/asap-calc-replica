"""Explicit period/runtime input container; no hidden date or period inference."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from types import MappingProxyType
from typing import Mapping


@dataclass(frozen=True, slots=True)
class PeriodContext:
    deal_id: str
    period: int
    payment_date: date | None = None
    runtime_inputs: Mapping[str, object] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.deal_id.strip():
            raise ValueError("deal_id must not be blank")
        if isinstance(self.period, bool) or not isinstance(self.period, int):
            raise TypeError("period must be an explicitly supplied integer")
        object.__setattr__(
            self, "runtime_inputs", MappingProxyType(dict(self.runtime_inputs))
        )

    def require_input(self, name: str) -> object:
        if name not in self.runtime_inputs:
            raise KeyError(f"Missing explicit runtime input: {name}")
        return self.runtime_inputs[name]
