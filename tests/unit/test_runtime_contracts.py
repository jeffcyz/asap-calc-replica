"""Tests for the fail-closed ASAP runtime contract boundaries."""

from datetime import date

import pytest

from asap_calc_replica.runtime import (
    DependencyKind, DependencyRegistry, PeriodContext, RuntimeContract,
)
from asap_calc_replica.runtime.collection_semantics import (
    UnresolvedCollectionSemantics, require_documented_collection,
)


def test_duplicate_registration_is_rejected():
    contract = RuntimeContract("test.method", DependencyKind.RUNTIME_METHOD)
    with pytest.raises(ValueError, match="Duplicate"):
        DependencyRegistry([contract, contract])


def test_collection_behavior_cannot_be_guessed():
    registry = DependencyRegistry([
        RuntimeContract("assets.synthetic", DependencyKind.COLLECTION)
    ])
    with pytest.raises(UnresolvedCollectionSemantics):
        require_documented_collection("assets.synthetic", registry)


def test_documented_collection_contract_is_returned_not_executed():
    contract = RuntimeContract(
        "assets.synthetic",
        DependencyKind.COLLECTION,
        semantics_documented=True,
        evidence_reference="synthetic-test-spec",
        side_effects_resolved=True,
    )
    assert require_documented_collection(
        "assets.synthetic", DependencyRegistry([contract])
    ) is contract


def test_period_context_requires_explicit_inputs():
    inputs = {"SyntheticRate": 0.05}
    context = PeriodContext(
        "AALC0000", 19, date(2026, 9, 15), inputs
    )
    inputs["SyntheticRate"] = 0.75
    assert context.require_input("SyntheticRate") == 0.05
    with pytest.raises(KeyError):
        context.require_input("UnknownIndex")


def test_bool_period_is_not_accepted():
    with pytest.raises(TypeError):
        PeriodContext("AALC0000", True)
