from __future__ import annotations

from typing import Any, Mapping

from emc_institutional_model.params import ModelParams


def project_pack_to_model_params_hints(payload: Mapping[str, Any]) -> dict[str, Any]:
    """
    Derive **partial** hints for building `ModelParams`. Always merge with explicit defaults;
    this does not replace a full financial model run.
    """
    scale = payload.get("scale") or {}
    ct = payload.get("capex_triplet") or {}
    lights = int(scale.get("number_of_lights") or 0)
    capex_ours = float(ct.get("capex_ours") or 0.0)
    per_light = (capex_ours / lights) if lights else None

    return {
        "number_of_lights": lights or ModelParams.model_fields["number_of_lights"].default,
        "number_of_poles": int(scale.get("number_of_poles") or lights or ModelParams.model_fields["number_of_poles"].default),
        "notes": {
            "implied_capex_per_light_usd_from_pack": per_light,
            "source": "project_capex_pack.capex_triplet.capex_ours / scale.number_of_lights",
        },
    }
