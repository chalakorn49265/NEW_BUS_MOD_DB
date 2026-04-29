from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class EvidenceCard:
    title: str
    why: str
    evidence: str
    notes: str = ""


def default_evidence_cards() -> list[EvidenceCard]:
    """
    Keep this lightweight and pitch-friendly.

    We intentionally avoid hard-coding Mozambique-specific web claims here; the dashboard can
    show these as placeholders and sales can swap links for the actual project.
    """
    return [
        EvidenceCard(
            title="Electricity bill goes to ~0 (solar)",
            why="AI+Solar system is designed to generate its own energy, so grid electricity costs are avoided for lighting.",
            evidence="(Add local project spec / design doc / measured PV output assumptions)",
        ),
        EvidenceCard(
            title="Single subscription replaces multiple uncertain costs",
            why="Client replaces electricity + maintenance variability with one predictable subscription.",
            evidence="(Add draft contract term sheet / proposal)",
        ),
        EvidenceCard(
            title="Every dollar is accounted for",
            why="We show exactly how subscription cashflows split into government fees, intermediaries, and provider O&M and margin.",
            evidence="(Generated audit workbook from this dashboard)",
        ),
    ]

