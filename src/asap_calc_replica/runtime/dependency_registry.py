"""A registry of reviewed runtime contracts; no undocumented auto-inference."""

from __future__ import annotations

from collections.abc import Iterable

from .object_contracts import RuntimeContract


class DependencyRegistry:
    def __init__(self, contracts: Iterable[RuntimeContract] = ()) -> None:
        self._contracts: dict[str, RuntimeContract] = {}
        for contract in contracts:
            self.register(contract)

    def register(self, contract: RuntimeContract) -> None:
        if contract.symbol in self._contracts:
            raise ValueError(f"Duplicate runtime dependency symbol: {contract.symbol}")
        self._contracts[contract.symbol] = contract

    def get(self, symbol: str) -> RuntimeContract | None:
        return self._contracts.get(symbol)

    def require(self, symbol: str) -> RuntimeContract:
        contract = self.get(symbol)
        if contract is None:
            raise KeyError(f"No documented contract registered for: {symbol}")
        return contract
