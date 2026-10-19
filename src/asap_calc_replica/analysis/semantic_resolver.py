"""Conservative A/B/C/D calculation reproducibility triage.

This is a metadata assessor, NOT an ASAP script parser or working formula
compiler. The calling analysis must provide verified field expressions,
runtime contracts, available input names and translation evidence.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from collections.abc import Collection

from asap_calc_replica.runtime.dependency_registry import DependencyRegistry


class Reproducibility(str, Enum):
    A = "A"  # Directly reproducible with complete, verified local inputs
    B = "B"  # Requires documented ASAP runtime mapping; inputs available
    C = "C"  # Understood semantics, but one or more inputs are missing
    D = "D"  # Unknown or unverified semantics / translation equivalence


@dataclass(frozen=True, slots=True)
class FieldEvidence:
    output_field: str
    dependencies: tuple[str, ...] = ()
    expression_semantics_documented: bool = False
    excel_equivalence_supported: bool = False


@dataclass(frozen=True, slots=True)
class FieldAssessment:
    output_field: str
    level: Reproducibility
    missing_inputs: tuple[str, ...] = ()
    unresolved_semantics: tuple[str, ...] = ()
    evidence_references: tuple[str, ...] = ()


def assess_field(
    field: FieldEvidence,
    registry: DependencyRegistry,
    available_inputs: Collection[str],
) -> FieldAssessment:
    """Classify using explicit metadata; D takes precedence over C over B/A.

    An A/B assessment is only a *readiness assessment*, not a validation
    assertion. Agreement with ASAP SQL still requires Excel recalculation.
    """
    if not field.output_field.strip():
        raise ValueError("output_field must not be blank")

    unresolved: set[str] = set()
    missing: set[str] = set()
    evidence: set[str] = set()
    uses_runtime_mapping = False
    present = set(available_inputs)

    if not field.expression_semantics_documented:
        unresolved.add("field expression semantics")
    if not field.excel_equivalence_supported:
        unresolved.add("Excel translation equivalence")

    for symbol in field.dependencies:
        contract = registry.get(symbol)
        if contract is None:
            unresolved.add(symbol)
            continue
        if not contract.is_semantically_resolved:
            unresolved.add(symbol)
        if contract.evidence_reference:
            evidence.add(contract.evidence_reference)
        uses_runtime_mapping |= contract.requires_runtime_mapping
        missing.update(name for name in contract.required_inputs if name not in present)

    if unresolved:
        level = Reproducibility.D
    elif missing:
        level = Reproducibility.C
    elif uses_runtime_mapping:
        level = Reproducibility.B
    else:
        level = Reproducibility.A

    return FieldAssessment(
        output_field=field.output_field,
        level=level,
        missing_inputs=tuple(sorted(missing)),
        unresolved_semantics=tuple(sorted(unresolved)),
        evidence_references=tuple(sorted(evidence)),
    )
