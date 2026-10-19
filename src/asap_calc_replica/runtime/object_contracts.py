"""Evidence-backed declarations of ASAP runtime dependencies.

This module does not assert behavior of any real ASAP method. A contract must
be documented from an authorized source and explicitly reviewed before it can
be treated as resolved by the semantic assessment layer.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class DependencyKind(str, Enum):
    SOURCE_FIELD = "source_field"
    COLLECTION = "collection"
    RUNTIME_METHOD = "runtime_method"
    RUNTIME_PARAMETER = "runtime_parameter"
    HISTORICAL_STATE = "historical_state"
    EXTERNAL_VALUE = "external_value"
    EXECUTION_LIFECYCLE = "execution_lifecycle"


@dataclass(frozen=True, slots=True)
class RuntimeContract:
    symbol: str
    kind: DependencyKind
    required_inputs: tuple[str, ...] = ()
    semantics_documented: bool = False
    evidence_reference: str | None = None
    side_effects_resolved: bool = False
    requires_runtime_mapping: bool = False
    runtime_mapping_documented: bool = False

    def __post_init__(self) -> None:
        if not self.symbol.strip():
            raise ValueError("A runtime dependency symbol must be nonempty")
        if any(not name.strip() for name in self.required_inputs):
            raise ValueError("Required input identifiers must be nonempty")

    @property
    def is_semantically_resolved(self) -> bool:
        return (
            self.semantics_documented
            and bool(self.evidence_reference and self.evidence_reference.strip())
            and self.side_effects_resolved
            and (not self.requires_runtime_mapping or self.runtime_mapping_documented)
        )
