"""EMC institutional street-lighting financial engine (MOZ v2 parity + TOU + MC)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    # Import only for typing; at runtime we lazily import to avoid hard deps (e.g. pandas)
    # when users only need submodules like `emc_institutional_model.laas`.
    from emc_institutional_model.runner import ModelResult as ModelResult
    from emc_institutional_model.runner import run_model as run_model

__all__ = ["run_model", "ModelResult"]


def __getattr__(name: str) -> Any:
    if name in {"run_model", "ModelResult"}:
        from emc_institutional_model.runner import ModelResult, run_model

        return {"run_model": run_model, "ModelResult": ModelResult}[name]
    raise AttributeError(name)
