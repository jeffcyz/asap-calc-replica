"""Fail-closed gating for collection-method semantics.

No active-asset selection rule is encoded here. The actual rules and
execution timing must come from authorized ASAP documentation or tests.
"""

from __future__ import annotations

from .dependency_registry import DependencyRegistry
from .object_contracts import DependencyKind, RuntimeContract


class UnresolvedCollectionSemantics(ValueError):
    pass


def require_documented_collection(
    symbol: str, registry: DependencyRegistry
) -> RuntimeContract:
    contract = registry.get(symbol)
    if contract is None or contract.kind != DependencyKind.COLLECTION:
        raise UnresolvedCollectionSemantics(
            f"Missing collection contract for {symbol!r}"
        )
    if not contract.is_semantically_resolved:
        raise UnresolvedCollectionSemantics(
            f"Collection behavior is not fully documented: {symbol!r}"
        )
    return contract
