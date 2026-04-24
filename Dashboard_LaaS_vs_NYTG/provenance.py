from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Mapping, Sequence


@dataclass(frozen=True)
class SourceRef:
    file: str
    sheet: str | None = None
    row_label: str | None = None
    notes: str | None = None


@dataclass(frozen=True)
class Provenance:
    """Explains where a number came from and how it was derived.

    This is intended for dashboard drill-down: nothing should be a mystery.
    """

    sources: tuple[SourceRef, ...]
    units: str
    transform: str

    def as_dict(self) -> dict[str, Any]:
        return {
            "units": self.units,
            "transform": self.transform,
            "sources": [
                {
                    "file": s.file,
                    "sheet": s.sheet,
                    "row_label": s.row_label,
                    "notes": s.notes,
                }
                for s in self.sources
            ],
        }


@dataclass(frozen=True)
class SeriesWithProv:
    """A year-indexed series with provenance."""

    values_by_year: dict[int, float]
    provenance: Provenance

    def years(self) -> list[int]:
        return sorted(self.values_by_year.keys())

    def get(self, year: int, default: float = 0.0) -> float:
        return float(self.values_by_year.get(int(year), default))

    def reindex_years(self, years: Sequence[int], fill: float = 0.0) -> SeriesWithProv:
        return SeriesWithProv(
            values_by_year={int(y): float(self.values_by_year.get(int(y), fill)) for y in years},
            provenance=self.provenance,
        )


def sum_series(series: Iterable[SeriesWithProv], *, units: str, transform: str, sources: Sequence[SourceRef]) -> SeriesWithProv:
    vals: dict[int, float] = {}
    for s in series:
        for y, v in s.values_by_year.items():
            vals[int(y)] = float(vals.get(int(y), 0.0)) + float(v)
    return SeriesWithProv(values_by_year=vals, provenance=Provenance(tuple(sources), units=units, transform=transform))


def dict_to_table(rows: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    """Utility for Streamlit tables; keeps plain-JSON serializable."""
    return [dict(r) for r in rows]

