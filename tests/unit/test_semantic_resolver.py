"""Synthetic contract tests: no real ASAP source code or client data."""

from asap_calc_replica.analysis.semantic_resolver import (
    FieldEvidence,
    Reproducibility,
    assess_field,
)
from asap_calc_replica.runtime import (
    DependencyKind,
    DependencyRegistry,
    RuntimeContract,
)


def _field(symbol: str) -> FieldEvidence:
    return FieldEvidence(
        "SyntheticCalculatedValue",
        (symbol,),
        expression_semantics_documented=True,
        excel_equivalence_supported=True,
    )


def _contract(symbol: str, *, mapped: bool = False) -> RuntimeContract:
    return RuntimeContract(
        symbol=symbol,
        kind=DependencyKind.RUNTIME_METHOD if mapped else DependencyKind.SOURCE_FIELD,
        required_inputs=("SyntheticInput",),
        semantics_documented=True,
        evidence_reference="synthetic-test-spec",
        side_effects_resolved=True,
        requires_runtime_mapping=mapped,
        runtime_mapping_documented=mapped,
    )


def test_direct_with_available_input_is_a():
    registry = DependencyRegistry([_contract("row.value")])
    result = assess_field(_field("row.value"), registry, {"SyntheticInput"})
    assert result.level == Reproducibility.A


def test_documented_runtime_mapping_is_b():
    registry = DependencyRegistry([_contract("service.value", mapped=True)])
    result = assess_field(_field("service.value"), registry, {"SyntheticInput"})
    assert result.level == Reproducibility.B


def test_missing_input_is_c():
    registry = DependencyRegistry([_contract("row.value")])
    result = assess_field(_field("row.value"), registry, set())
    assert result.level == Reproducibility.C
    assert result.missing_inputs == ("SyntheticInput",)


def test_undocumented_or_unknown_semantics_is_d_even_when_data_is_missing():
    registry = DependencyRegistry([RuntimeContract("unknown.call", DependencyKind.RUNTIME_METHOD)])
    result = assess_field(_field("unknown.call"), registry, set())
    assert result.level == Reproducibility.D
    assert result.unresolved_semantics == ("unknown.call",)


def test_missing_contract_is_d():
    result = assess_field(_field("not.registered"), DependencyRegistry(), set())
    assert result.level == Reproducibility.D


def test_translation_without_evidence_is_d():
    registry = DependencyRegistry([_contract("row.value")])
    field = FieldEvidence("SyntheticCalculatedValue", ("row.value",))
    assert assess_field(field, registry, {"SyntheticInput"}).level == Reproducibility.D
