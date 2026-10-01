import io
import datetime
from typing import Dict, List, Any, Optional

import reportlab
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas


class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas to dynamically compute and print exact 'Página X de Y',
    plus standard running header on secondary pages and running footer on all pages.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_decorations(self, page_count):
        self.saveState()
        page_width, page_height = landscape(A4)

        # Header for page 2 and onwards
        if self._pageNumber > 1:
            self.setFont("Helvetica-Bold", 8)
            self.setFillColor(colors.HexColor("#1E3A8A"))
            self.drawString(36, page_height - 25, "RELATÓRIO EXECUTIVO DE GESTÃO DE DORES — ANÁLISE OPERACIONAL & FINANCEIRA")
            
            self.setFont("Helvetica", 7.5)
            self.setFillColor(colors.HexColor("#64748B"))
            period_text = getattr(self, "period_label", "Todo o Histórico")
            self.drawRightString(page_width - 36, page_height - 25, f"Período: {period_text}")
            
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(36, page_height - 30, page_width - 36, page_height - 30)

        # Running Footer on all pages
        self.setStrokeColor(colors.HexColor("#E2E8F0"))
        self.setLineWidth(0.5)
        self.line(36, 28, page_width - 36, 28)

        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#64748B"))
        gen_time = getattr(self, "generation_time", datetime.datetime.now().strftime("%d/%m/%Y às %H:%M"))
        footer_left = f"Mapeamento de Dores | Base: Microsoft Excel | Gerado em: {gen_time} | Documento Confidencial"
        self.drawString(36, 18, footer_left)

        page_str = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(page_width - 36, 18, page_str)
        self.restoreState()


def format_currency_br(val: Optional[float]) -> str:
    if val is None or val == "":
        return "R$ 0,00"
    try:
        val = float(val)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"


def format_kg_br(val: Optional[float]) -> str:
    if val is None or val == "":
        return "0 kg"
    try:
        val = float(val)
        if val.is_integer():
            return f"{int(val):,}".replace(",", ".") + " kg"
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"
    except (ValueError, TypeError):
        return "0 kg"


def generate_executive_insights(records: List[Dict[str, Any]], total_valor: float, total_kg: float) -> Dict[str, str]:
    """
    Synthesizes intelligent, executive-level diagnostic insights across:
    1. Problems & Root Causes
    2. Financial & Accounting Impact
    3. Physical Stock & Cargo Volume
    4. Action Plans & SLA Compliance
    """
    total = len(records)
    if total == 0:
        return {
            "problemas": "Nenhuma ocorrência registrada no período selecionado.",
            "financeiro": "Sem impacto financeiro no período.",
            "estoque": "Sem impacto físico em estoque no período.",
            "tratativas": "Sem tratativas ou planos de ação pendentes."
        }

    # 1. Categorias & Causas
    by_cat = {}
    by_setor_resp = {}
    by_filial = {}
    recorrentes_count = 0
    com_plano_count = 0
    concluidas_count = 0
    vencidos_count = 0

    for r in records:
        c = r.get("categoria_dor") or "Não Informada"
        by_cat[c] = by_cat.get(c, 0) + 1

        s = r.get("setor_responsavel") or "Não Informado"
        by_setor_resp[s] = by_setor_resp.get(s, 0) + 1

        f = r.get("filial") or "Não Informada"
        by_filial[f] = by_filial.get(f, 0) + 1

        if str(r.get("recorrente", "")).strip().lower() in ["sim", "s", "true"]:
            recorrentes_count += 1

        if r.get("plano_acao") and str(r.get("plano_acao")).strip():
            com_plano_count += 1

        if r.get("status") == "Concluído":
            concluidas_count += 1

        if r.get("situacao_sla") == "Vencido":
            vencidos_count += 1

    top_cat = max(by_cat.items(), key=lambda x: x[1])[0] if by_cat else "N/A"
    top_cat_pct = round((by_cat.get(top_cat, 0) / total) * 100, 1) if total > 0 else 0

    top_setor = max(by_setor_resp.items(), key=lambda x: x[1])[0] if by_setor_resp else "N/A"
    top_filial = max(by_filial.items(), key=lambda x: x[1])[0] if by_filial else "N/A"

    pct_recorrente = round((recorrentes_count / total) * 100, 1)
    pct_plano = round((com_plano_count / total) * 100, 1)
    taxa_resolucao = round((concluidas_count / total) * 100, 1)
    ticket_medio = total_valor / total if total > 0 else 0

    # Top financial record
    sorted_fin = sorted(records, key=lambda x: float(x.get("valor_notas") or 0.0), reverse=True)
    top1_fin = sorted_fin[0] if sorted_fin and float(sorted_fin[0].get("valor_notas") or 0.0) > 0 else None

    # Text 1: Problemas & Causas Raízes
    insight_prob = (
        f"<b>Concentração Crítica:</b> A categoria <b>{top_cat}</b> lidera as ocorrências no período, "
        f"representando <b>{top_cat_pct}%</b> do total avaliado ({by_cat.get(top_cat, 0)} caso(s)). "
        f"O setor com maior demanda de atendimento atribuída é <b>{top_setor}</b>. "
    )
    if pct_recorrente > 0:
        insight_prob += (
            f"Alerta: <b>{pct_recorrente}%</b> das dores ({recorrentes_count} ocorrência(s)) foram classificadas como "
            f"<b>Recorrentes</b>, evidenciando falhas processuais crônicas ou necessidade de parametrização sistêmica "
            f"para evitar reincidência de retrabalho."
        )
    else:
        insight_prob += "As dores registradas configuram eventos pontuais, sem histórico crítico de reincidência imediata."

    # Text 2: Impactos Financeiros & Contábil
    insight_fin = (
        f"<b>Exposição Financeira:</b> O montante acumulado nas ocorrências do período totaliza <b>{format_currency_br(total_valor)}</b>, "
        f"com impacto médio de <b>{format_currency_br(ticket_medio)}</b> por demanda. "
    )
    if top1_fin:
        insight_fin += (
            f"O maior evento financeiro isolado refere-se à demanda <b>#{top1_fin.get('id')}</b> na filial <b>{top1_fin.get('filial')}</b> "
            f"no valor de <b>{format_currency_br(top1_fin.get('valor_notas'))}</b> ({top1_fin.get('categoria_dor')}). "
        )
    insight_fin += (
        "<b>Implicações Contábeis:</b> Riscos associados a duplicidade de escrituração de notas, passivos fiscais, "
        "retenção de faturamento e contingências de frete/armazenagem exigem conciliação imediata entre Fiscal e Financeiro."
    )

    # Text 3: Estoque & Carga
    insight_est = (
        f"<b>Volume Físico Afetado:</b> O total de mercadorias envolvidas nas pendências do período atinge <b>{format_kg_br(total_kg)}</b>. "
        f"A unidade de maior concentração de carga afetada é <b>{top_filial}</b>. "
        f"<b>Impacto Operacional no Armazém:</b> Divergências físicas versus sistêmicas geram distorções no inventário contábil, "
        f"bloqueio preventivo de lotes para expedição e risco de atrasos na entrega aos clientes finais. "
        f"Recomenda-se acuracidade imediata de saldo físico."
    )

    # Text 4: Tratativas & SLA
    insight_trat = (
        f"<b>Aderência ao Plano de Ação:</b> <b>{pct_plano}%</b> das ocorrências ({com_plano_count} caso(s)) possuem plano de ação "
        f"estruturado. A taxa de conclusão geral no período é de <b>{taxa_resolucao}%</b> ({concluidas_count} concluída(s)). "
    )
    if vencidos_count > 0:
        insight_trat += (
            f"<font color='#DC2626'><b>Atenção Executiva:</b> Há <b>{vencidos_count} demanda(s) com SLA Vencido</b>, "
            f"requerendo escalonamento urgente junto às lideranças responsáveis.</font>"
        )
    else:
        insight_trat += "<font color='#059669'><b>Excelente Aderência:</b> 100% das demandas em aberto estão cumprindo o SLA estipulado.</font>"

    return {
        "problemas": insight_prob,
        "financeiro": insight_fin,
        "estoque": insight_est,
        "tratativas": insight_trat
    }


def build_pdf_report(
    records: List[Dict[str, Any]],
    data_inicio: Optional[str] = None,
    data_fim: Optional[str] = None,
    periodo_label: Optional[str] = None
) -> bytes:
    """
    Generates a complete, multi-page executive PDF report in landscape A4.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    # Styles
    styles = getSampleStyleSheet()

    # Custom typography
    style_title = ParagraphStyle(
        "ReportTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#FFFFFF")
    )

    style_subtitle = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#E2E8F0")
    )

    style_meta_box = ParagraphStyle(
        "MetaBox",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#F8FAFC"),
        alignment=2 # Right
    )

    style_kpi_num = ParagraphStyle(
        "KpiNum",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=15,
        alignment=1 # Center
    )

    style_kpi_label = ParagraphStyle(
        "KpiLabel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=9,
        textColor=colors.HexColor("#64748B"),
        alignment=1 # Center
    )

    style_kpi_sub = ParagraphStyle(
        "KpiSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#94A3B8"),
        alignment=1 # Center
    )

    style_section_heading = ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=4
    )

    style_box_header = ParagraphStyle(
        "BoxHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#FFFFFF")
    )

    style_box_body = ParagraphStyle(
        "BoxBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=10.5,
        textColor=colors.HexColor("#1E293B")
    )

    style_table_header = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=9,
        textColor=colors.HexColor("#FFFFFF"),
        alignment=1
    )

    style_table_cell = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.5,
        leading=8.5,
        textColor=colors.HexColor("#0F172A")
    )

    style_table_cell_bold = ParagraphStyle(
        "TableCellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=8.5,
        textColor=colors.HexColor("#0F172A")
    )

    style_table_cell_center = ParagraphStyle(
        "TableCellCenter",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.5,
        leading=8.5,
        alignment=1,
        textColor=colors.HexColor("#0F172A")
    )

    style_badge_concluido = ParagraphStyle(
        "BadgeConc", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6, leading=7, alignment=1, textColor=colors.HexColor("#166534")
    )
    style_badge_aberto = ParagraphStyle(
        "BadgeAberto", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6, leading=7, alignment=1, textColor=colors.HexColor("#92400E")
    )
    style_badge_vencido = ParagraphStyle(
        "BadgeVencido", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6, leading=7, alignment=1, textColor=colors.HexColor("#991B1B")
    )

    # Computations
    total_records = len(records)
    total_valor = sum(float(r.get("valor_notas") or 0.0) for r in records)
    total_kg = sum(float(r.get("qtd_nota_kg") or 0.0) for r in records if r.get("qtd_nota_kg") is not None)
    concluidas = sum(1 for r in records if r.get("status") == "Concluído")
    abertas = sum(1 for r in records if r.get("status") not in ["Concluído", "Cancelado"])
    sla_vencido = sum(1 for r in records if r.get("situacao_sla") == "Vencido")
    sla_avencer = sum(1 for r in records if r.get("situacao_sla") == "A vencer")
    sla_noprazo = sum(1 for r in records if r.get("situacao_sla") == "No prazo")
    taxa_resolucao = round((concluidas / total_records * 100), 1) if total_records > 0 else 0

    # Period display
    if not periodo_label:
        if data_inicio and data_fim:
            def format_d(d_str):
                try:
                    parts = d_str.split("-")
                    return f"{parts[2]}/{parts[1]}/{parts[0]}"
                except Exception:
                    return d_str
            periodo_label = f"De {format_d(data_inicio)} até {format_d(data_fim)}"
        elif data_inicio:
            periodo_label = f"A partir de {data_inicio}"
        elif data_fim:
            periodo_label = f"Até {data_fim}"
        else:
            periodo_label = "Todo o Histórico"

    now_str = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")

    story = []

    # =========================================================================
    # 1. TOP HEADER BANNER
    # =========================================================================
    banner_left = [
        Paragraph("RELATÓRIO EXECUTIVO DE MAPEAMENTO DE DORES", style_title),
        Spacer(1, 2),
        Paragraph("Diagnóstico Operacional, Tratativas, Impactos Financeiros, Contábeis e Estoque", style_subtitle)
    ]
    banner_right = [
        Paragraph(f"<b>Período Analisado:</b> {periodo_label}", style_meta_box),
        Paragraph(f"<b>Base de Dados:</b> Microsoft Excel Integrado", style_meta_box),
        Paragraph(f"<b>Emissão:</b> {now_str}", style_meta_box)
    ]

    header_table = Table([[banner_left, banner_right]], colWidths=[520, 249.89])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1E3A8A")),
        ("PADDING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    # =========================================================================
    # 2. EXECUTIVE KPI CARDS (6 CARDS)
    # =========================================================================
    card1 = [
        Paragraph("TOTAL DE DORES", style_kpi_label),
        Paragraph(f"<font color='#1E3A8A'>{total_records}</font>", style_kpi_num),
        Paragraph("Ocorrências no período", style_kpi_sub)
    ]
    card2 = [
        Paragraph("RESOLUÇÃO", style_kpi_label),
        Paragraph(f"<font color='#166534'>{taxa_resolucao}%</font>", style_kpi_num),
        Paragraph(f"{concluidas} conc. | {abertas} abertas", style_kpi_sub)
    ]
    sla_color = "#DC2626" if sla_vencido > 0 else "#166534"
    card3 = [
        Paragraph("SITUAÇÃO DO SLA", style_kpi_label),
        Paragraph(f"<font color='{sla_color}'>{sla_vencido} Vencido(s)</font>", style_kpi_num),
        Paragraph(f"{sla_noprazo} no prazo | {sla_avencer} a vencer", style_kpi_sub)
    ]
    card4 = [
        Paragraph("IMPACTO FINANCEIRO", style_kpi_label),
        Paragraph(f"<font color='#059669'>{format_currency_br(total_valor)}</font>", style_kpi_num),
        Paragraph("Valor contábil envolvido", style_kpi_sub)
    ]
    card5 = [
        Paragraph("ESTOQUE / CARGA", style_kpi_label),
        Paragraph(f"<font color='#2563EB'>{format_kg_br(total_kg)}</font>", style_kpi_num),
        Paragraph("Volume físico impactado", style_kpi_sub)
    ]
    total_anexos_count = sum(len(r.get("anexos") or []) for r in records)
    card6 = [
        Paragraph("EVIDÊNCIAS & ANEXOS", style_kpi_label),
        Paragraph(f"<font color='#4F46E5'>{total_anexos_count}</font>", style_kpi_num),
        Paragraph("Documentos e arquivos", style_kpi_sub)
    ]

    kpi_col_w = 769.89 / 6.0
    kpi_table = Table([[card1, card2, card3, card4, card5, card6]], colWidths=[kpi_col_w]*6)
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("PADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 12))

    # =========================================================================
    # 3. EXECUTIVE STRATEGIC INSIGHTS (2x2 GRID)
    # =========================================================================
    insights = generate_executive_insights(records, total_valor, total_kg)

    story.append(Paragraph("PAINEL DE INSIGHTS ESTRATÉGICOS & IMPACTOS OPERACIONAIS", style_section_heading))
    story.append(Spacer(1, 4))

    def make_insight_box(title: str, text: str, header_bg: str, border_color: str, width: float):
        box_header = Table([[Paragraph(title, style_box_header)]], colWidths=[width])
        box_header.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(header_bg)),
            ("PADDING", (0, 0), (-1, -1), 5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))

        box_body = Table([[Paragraph(text, style_box_body)]], colWidths=[width])
        box_body.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(border_color)),
            ("PADDING", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))

        return [box_header, box_body]

    half_w = (769.89 - 10) / 2.0
    box_prob = make_insight_box(
        "1. DIAGNÓSTICO DAS DORES & CAUSAS RAÍZES",
        insights["problemas"],
        "#1E3A8A", "#93C5FD", half_w
    )
    box_fin = make_insight_box(
        "2. IMPACTO FINANCEIRO & IMPLICAÇÕES CONTÁBEIS",
        insights["financeiro"],
        "#047857", "#6EE7B7", half_w
    )
    box_est = make_insight_box(
        "3. IMPACTO FÍSICO EM ESTOQUE & MOVIMENTAÇÃO (KG)",
        insights["estoque"],
        "#B45309", "#FCD34D", half_w
    )
    box_trat = make_insight_box(
        "4. EFICÁCIA DAS TRATATIVAS & CUMPRIMENTO DE SLA",
        insights["tratativas"],
        "#4338CA", "#A5B4FC", half_w
    )

    insights_grid = Table([
        [box_prob, box_fin],
        [box_est, box_trat]
    ], colWidths=[half_w + 5, half_w + 5])
    insights_grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("PADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
    ]))
    story.append(insights_grid)

    # =========================================================================
    # 4. PAGE BREAK TO DETAILED OPERATIONAL ROSTER
    # =========================================================================
    story.append(PageBreak())

    story.append(Paragraph("QUADRO ANALÍTICO DE OCORRÊNCIAS, TRATATIVAS E IMPACTOS", style_section_heading))
    story.append(Paragraph(
        "Listagem detalhada das ocorrências registradas no período selecionado, com valores contábeis, peso físico e situação do SLA.",
        style_subtitle
    ))
    story.append(Spacer(1, 6))

    # Detailed Table
    col_w = [
        26,   # ID
        48,   # Data Reg
        70,   # Filial / Cat
        68,   # Setor
        175,  # Descrição & Causa
        54,   # Qtd kg
        68,   # Valor R$
        60,   # Prio / SLA
        50,   # Status
        150.89 # Plano de Ação
    ]

    table_data = [[
        Paragraph("ID", style_table_header),
        Paragraph("Data", style_table_header),
        Paragraph("Filial / Categoria", style_table_header),
        Paragraph("Setor Responsável", style_table_header),
        Paragraph("Descrição do Problema & Causa Raiz", style_table_header),
        Paragraph("Estoque (kg)", style_table_header),
        Paragraph("Valor (R$)", style_table_header),
        Paragraph("Prioridade / SLA", style_table_header),
        Paragraph("Status", style_table_header),
        Paragraph("Plano de Ação / Solução", style_table_header),
    ]]

    for rec in records:
        rec_id = f"#{rec.get('id')}"
        dt = rec.get("data_registro_br") or "-"
        filial_cat = f"<b>{rec.get('filial', '-')}</b><br/><font color='#2563EB'>{rec.get('categoria_dor', '-')}</font>"
        setor = f"<b>{rec.get('setor_responsavel', '-')}</b><br/><font color='#64748B'>Resp: {rec.get('responsavel', '-')}</font>"

        desc_text = f"<b>{rec.get('descricao_problema', '-')}</b>"
        if rec.get("causa_raiz"):
            desc_text += f"<br/><font color='#475569'><i>Causa: {rec.get('causa_raiz')}</i></font>"
        if rec.get("nota_fiscal"):
            desc_text += f"<br/><font color='#6B7280'>NF {rec.get('nota_fiscal')}</font>"

        qtd_val = format_kg_br(rec.get("qtd_nota_kg"))
        val_contabil = format_currency_br(rec.get("valor_notas"))

        prio_sla = f"<b>{rec.get('prioridade', 'Média')}</b><br/>"
        sit_sla = rec.get("situacao_sla", "-")
        if sit_sla == "Vencido":
            prio_sla += "<font color='#DC2626'><b>Vencido</b></font>"
        elif sit_sla == "A vencer":
            prio_sla += "<font color='#D97706'><b>A vencer</b></font>"
        elif sit_sla == "Concluído":
            prio_sla += "<font color='#166534'>Concluído</font>"
        else:
            prio_sla += "<font color='#166534'>No prazo</font>"

        st_val = rec.get("status", "Aberto")
        if st_val == "Concluído":
            st_p = Paragraph(st_val, style_badge_concluido)
        elif sit_sla == "Vencido":
            st_p = Paragraph(st_val, style_badge_vencido)
        else:
            st_p = Paragraph(st_val, style_badge_aberto)

        plano_text = rec.get("plano_acao") or "<i>Sem plano de ação registrado.</i>"
        if rec.get("responsavel_solucao"):
            plano_text += f"<br/><font color='#475569'>Solução: {rec.get('responsavel_solucao')}</font>"

        table_data.append([
            Paragraph(rec_id, style_table_cell_center),
            Paragraph(dt, style_table_cell_center),
            Paragraph(filial_cat, style_table_cell),
            Paragraph(setor, style_table_cell),
            Paragraph(desc_text, style_table_cell),
            Paragraph(f"<b>{qtd_val}</b>", style_table_cell_center),
            Paragraph(f"<b>{val_contabil}</b>", style_table_cell_center),
            Paragraph(prio_sla, style_table_cell_center),
            st_p,
            Paragraph(plano_text, style_table_cell)
        ])

    roster_table = Table(table_data, colWidths=col_w, repeatRows=1)
    roster_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#FFFFFF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
        ("PADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(roster_table)

    # Build PDF with custom NumberedCanvas
    canvas_maker = NumberedCanvas
    # Inject period label and generation time into canvas
    NumberedCanvas.period_label = periodo_label
    NumberedCanvas.generation_time = now_str

    doc.build(story, canvasmaker=canvas_maker)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes
