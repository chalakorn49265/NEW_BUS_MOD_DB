from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class EvidenceCard:
    title: str
    why: str
    evidence_url: str
    maps_to_cells: list[str]
    notes: str = ""


def cards_for_selected_tier(*, has_upfront: bool, has_tail_discount: bool, laas_saving_rate: float | None) -> list[EvidenceCard]:
    cards: list[EvidenceCard] = []

    cards.append(
        EvidenceCard(
            title="智能调光/自适应照明 → 电费节省更高",
            why="通过按时段/车流人流自适应调光，减少不必要照明功率与用电量，电费下降。",
            evidence_url="https://mdpi-res.com/d_attachment/smartcities/smartcities-03-00071/article_deploy/smartcities-03-00071-v2.pdf",
            maps_to_cells=["02_Inputs!D31(Laas节电率)", "03_Baseline!D12(基准电费Y1)", "05_Annual_Model!D62:M62(业主年度净节省)"],
            notes=f"模型中节电率当前为 {laas_saving_rate:.1%}（如可读取）。" if isinstance(laas_saving_rate, float) else "节电率取自工作簿输入。",
        )
    )

    cards.append(
        EvidenceCard(
            title="预测性维护/故障自动识别 → 人工与材料与车辆成本降低",
            why="AI对路灯故障/异常用电进行自动识别与分派，减少人工巡检频次与出车次数，并降低备件浪费。",
            evidence_url="https://oxmaint.com/industries/government/street-lighting-lamp-out-ai-detection",
            maps_to_cells=["02_Inputs!D22(人工/维修按灯)", "02_Inputs!D23(平台费)", "02_Inputs!D24(备件)", "05_Annual_Model!D61:M61(业主年度总支出)"],
        )
    )

    if has_upfront:
        cards.append(
            EvidenceCard(
                title="首期款/预付 → 年费更低（摊销抵扣）",
                why="首期款作为预付在合同期内摊销抵扣年费，使后续年度订阅支出下降。",
                evidence_url="(商业条款/合同口径)",
                maps_to_cells=["02_Inputs!D45(首期款)", "05_Annual_Model!D40:M40(年费时间表)", "05_Annual_Model!D62:M62(净节省)"],
                notes="此项为商业条款假设/谈判变量，不是外部研究结论。",
            )
        )

    if has_tail_discount:
        cards.append(
            EvidenceCard(
                title="后4年减免 → 中后期现金流更友好（不降到0）",
                why="在第7-10年对年费进行减免，形成阶梯式价格；同时保持底价，避免不合理归零。",
                evidence_url="(商业条款/合同口径)",
                maps_to_cells=["02_Inputs!D46(后4年减免)", "05_Annual_Model!D40:M40(年费时间表)"],
                notes="此项为商业条款假设/谈判变量。",
            )
        )

    cards.append(
        EvidenceCard(
            title="平台化运维 → 交付可量化（SLA）与管理效率提升",
            why="将运维流程平台化，形成可追踪工单、响应时间与可用率指标，改善管理效率与服务稳定性。",
            evidence_url="(行业常识/可补充公司案例或第三方白皮书)",
            maps_to_cells=["02_Inputs!D23(平台费)", "01_Dashboard!C19/D19(NPV)", "01_Dashboard!C20/D20(IRR)"],
            notes="若需要更强证据，可替换为指定白皮书/政府项目评估报告链接。",
        )
    )

    return cards

