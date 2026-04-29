from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


BaselineType = Literal["LED", "HPS"]


@dataclass(frozen=True)
class BaselineTech:
    """
    Explicit baseline-technology parameters (LED vs HPS).

    This is intentionally "sales-friendly" and transparent: baseline annual cost is electricity + maintenance.
    """

    name: str
    watt_per_light: float
    maintenance_usd_per_light_year: float


DEFAULT_BASELINES: dict[BaselineType, BaselineTech] = {
    # Defaults are placeholders; sales should override with local inventory when known.
    "LED": BaselineTech(name="Normal LED", watt_per_light=120.0, maintenance_usd_per_light_year=8.0),
    "HPS": BaselineTech(name="Old High-Pressure Sodium (HPS/高压钠灯)", watt_per_light=250.0, maintenance_usd_per_light_year=18.0),
}


def baseline_electricity_annual_usd(
    *,
    number_of_lights: int,
    watt_per_light: float,
    operating_hours_per_night: float,
    days_per_year: float,
    electricity_price_usd_per_kwh: float,
) -> float:
    kw = float(watt_per_light) / 1000.0
    kwh = float(number_of_lights) * kw * float(operating_hours_per_night) * float(days_per_year)
    return float(kwh) * float(electricity_price_usd_per_kwh)


def baseline_maintenance_annual_usd(*, number_of_lights: int, maintenance_usd_per_light_year: float) -> float:
    return float(number_of_lights) * float(maintenance_usd_per_light_year)

