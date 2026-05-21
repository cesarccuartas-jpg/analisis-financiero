from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ─── Márgenes de página ───────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.2)
    section.right_margin  = Cm(2.2)

# ─── Paleta de colores ────────────────────────────────────────────────────────
AZUL_OSC   = RGBColor(0x1F, 0x49, 0x7D)   # encabezados principales
AZUL_MED   = RGBColor(0x2E, 0x75, 0xB6)   # encabezados secundarios
AZUL_CLARO = RGBColor(0xBD, 0xD7, 0xEE)   # fondo tabla header
GRIS_CLARO = RGBColor(0xF2, 0xF2, 0xF2)   # filas alternas
ROJO       = RGBColor(0xC0, 0x00, 0x00)   # alertas / negativos
VERDE      = RGBColor(0x37, 0x86, 0x10)   # positivos
NARANJA    = RGBColor(0xC5, 0x5A, 0x11)   # advertencias
NEGRO      = RGBColor(0x00, 0x00, 0x00)
BLANCO     = RGBColor(0xFF, 0xFF, 0xFF)

# ─── Helpers ─────────────────────────────────────────────────────────────────
def set_cell_bg(cell, color: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    hex_color = '{:02X}{:02X}{:02X}'.format(color[0], color[1], color[2])
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, border_color='2E75B6', size=4):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),   'single')
        b.set(qn('w:sz'),    str(size))
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), border_color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def add_heading(text, level=1):
    p    = doc.add_paragraph()
    run  = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size  = Pt(16)
        run.font.color.rgb = AZUL_OSC
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after  = Pt(6)
        # línea inferior
        pPr  = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot  = OxmlElement('w:bottom')
        bot.set(qn('w:val'),   'single')
        bot.set(qn('w:sz'),    '8')
        bot.set(qn('w:space'), '1')
        bot.set(qn('w:color'), '1F497D')
        pBdr.append(bot)
        pPr.append(pBdr)
    elif level == 2:
        run.font.size  = Pt(13)
        run.font.color.rgb = AZUL_MED
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
    elif level == 3:
        run.font.size  = Pt(11)
        run.font.color.rgb = AZUL_OSC
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after  = Pt(3)
    return p

def add_body(text, bold=False, color=None, indent=0, space_after=4):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.bold = bold
    if color:
        run.font.color.rgb = color
    p.paragraph_format.left_indent  = Cm(indent * 0.5)
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    return p

def add_bullet(text, level=0, bold=False, color=None):
    p   = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.bold = bold
    if color:
        run.font.color.rgb = color
    p.paragraph_format.left_indent   = Cm(0.5 + level * 0.5)
    p.paragraph_format.space_after   = Pt(2)
    p.paragraph_format.space_before  = Pt(0)
    return p

def add_table(headers, rows, col_widths=None, alternate=True):
    n_cols = len(headers)
    table  = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Encabezado
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].font.size  = Pt(9)
        hdr_cells[i].paragraphs[0].runs[0].font.bold  = True
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = BLANCO
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(hdr_cells[i], AZUL_OSC)
        set_cell_borders(hdr_cells[i], 'FFFFFF', 4)

    # Filas de datos
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        bg = GRIS_CLARO if (alternate and r_idx % 2 == 1) else BLANCO
        for c_idx, cell_val in enumerate(row_data):
            cell = row_cells[c_idx]
            cell.text = str(cell_val)
            p   = cell.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run(str(cell_val))
            run.font.size = Pt(9)
            # Color especial para negativos/positivos
            v = str(cell_val)
            if v.startswith('-') and '$' in v:
                run.font.color.rgb = ROJO
            elif v.startswith('+') and '$' in v:
                run.font.color.rgb = VERDE
            else:
                run.font.color.rgb = NEGRO
            # Negrita en filas totales
            if any(kw in str(row_data[0]).upper() for kw in ['TOTAL','UTILIDAD','CMV','ACUMULADO']):
                run.bold = True
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_cell_bg(cell, bg)
            set_cell_borders(cell, '2E75B6', 4)

    # Anchos de columna
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)

    doc.add_paragraph()
    return table

def add_highlight_box(text, bg_color=None, text_color=BLANCO, bold=True):
    """Caja de texto destacada (tabla de 1 celda)."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.rows[0].cells[0]
    cell.text = text
    p   = cell.paragraphs[0]
    run = p.runs[0]
    run.font.size = Pt(10)
    run.font.bold = bold
    run.font.color.rgb = text_color
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_cell_bg(cell, bg_color or AZUL_OSC)
    set_cell_borders(cell, 'FFFFFF', 0)
    doc.add_paragraph()

def page_break():
    doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PORTADA
# ══════════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(60)
run = p.add_run('DIAGNÓSTICO FINANCIERO INTEGRAL')
run.font.size  = Pt(24)
run.font.bold  = True
run.font.color.rgb = AZUL_OSC
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
run = p.add_run('Empresa Metalmecánica')
run.font.size  = Pt(16)
run.font.color.rgb = AZUL_MED
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
run = p.add_run('Enero – Abril 2026 vs 2025')
run.font.size  = Pt(14)
run.font.color.rgb = AZUL_MED
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Cifras en millones de pesos colombianos (COP $M)')
run.font.size  = Pt(11)
run.font.color.rgb = RGBColor(0x59, 0x56, 0x59)
run.font.italic = True
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Preparado con base en:\n• 1-Relacion ventas 2026-1.xlsx\n• 2-Estado de resultados 2026-abril-completo.xlsm\n• 5-Informe costos 2024-2025-2026 abril.xlsx')
run.font.size  = Pt(10)
run.font.color.rgb = RGBColor(0x59, 0x56, 0x59)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Mayo 2026')
run.font.size  = Pt(11)
run.font.bold  = True
run.font.color.rgb = AZUL_OSC
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 1: RESUMEN EJECUTIVO
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 1: RESUMEN EJECUTIVO — 10 HALLAZGOS CLAVE', 1)

hallazgos = [
    ('1', 'Ventas crecen +11.3% (+$1,936M) pero utilidad neta cae -56.2% (-$2,086M).',
     'El crecimiento es real pero está siendo destruido por deterioro en márgenes brutos, explosión de costos de producción y un salto masivo en gastos familiares.', ROJO),
    ('2', 'Margen bruto cayó de 35.1% a 28.4% (-6.6 puntos porcentuales).',
     '$581M menos de utilidad bruta a pesar de vender $1,936M más. El crecimiento en ventas no solo no mejora la rentabilidad — la empeora.', ROJO),
    ('3', 'Las varillas pasaron del 1.6% al 11.8% de las ventas, con margen del 9.1%.',
     'Generaron $2,251M en ventas pero solo $204M de utilidad bruta. Están distorsionando completamente el mix y destruyendo el margen consolidado.', ROJO),
    ('4', 'El Fabricado Plástico colapsó su margen: de 32.8% a 18.4% (-14.4 pp).',
     'Con ventas de $2,315M, esto implica $262M de utilidad bruta destruida.', ROJO),
    ('5', 'Los costos de producción subieron +25.9% (+$913M), triplicando el crecimiento de los fabricados.',
     'Arriendos: +315% (+$254M). CIF: +52.3%. MOI: +24.5%. MOD: +14.8%. La planta absorbe más costo sin crecer en volumen de fabricados (metalmecánica cayó -2.8%).', NARANJA),
    ('6', 'Los gastos de la familia se dispararon +311% (+$781M): de $251M a $1,032M en 4 meses.',
     '"De personal familia": $968M. Este renglón pasó del 1.5% al 5.4% de las ventas. Es el ítem individual que más destruyó la utilidad neta.', ROJO),
    ('7', 'CE-125 y Adaptadores venden por debajo de costo.',
     'CE-125: pérdida bruta -$16M (margen -8.3%). Adaptadores: pérdida bruta -$6M (margen -4.7%). Dos líneas que consumen recursos y generan pérdida.', ROJO),
    ('8', 'La metalmecánica fabricada cayó en ventas (-2.8%, -$374M) mientras sus costos subieron +25.9%.',
     'Menor volumen + mayor costo fijo = deterioro de absorción y margen por unidad.', NARANJA),
    ('9', 'Marzo 2026 registró utilidad neta negativa: -$151M.',
     'Arrastrado por costos no operacionales de $504M (incluyendo $408M de intereses) y gastos familiares de $339M en ese solo mes.', ROJO),
    ('10', 'La utilidad operacional cayó de 22.7% a 15.3% de ventas (-7.4 pp, -$966M).',
     'Incluso antes de los efectos financieros y familiares, la operación ya genera significativamente menos valor.', NARANJA),
]

for num, titulo, detalle, color in hallazgos:
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    c0 = table.rows[0].cells[0]
    c1 = table.rows[0].cells[1]
    c0.text = num
    r0 = c0.paragraphs[0].runs[0]
    r0.font.size = Pt(14); r0.font.bold = True; r0.font.color.rgb = BLANCO
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_bg(c0, color)
    c0.width = Cm(1.0)
    c1.paragraphs[0].clear()
    r1 = c1.paragraphs[0].add_run(titulo + ' ')
    r1.font.size = Pt(10); r1.font.bold = True; r1.font.color.rgb = color
    c1.add_paragraph(detalle).runs[0].font.size = Pt(9)
    set_cell_bg(c1, GRIS_CLARO if int(num) % 2 == 0 else BLANCO)
    set_cell_borders(c0, 'FFFFFF', 2); set_cell_borders(c1, 'CCCCCC', 2)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 2: ANÁLISIS CUANTITATIVO
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 2: ANÁLISIS CUANTITATIVO', 1)

# 2.1 Ventas y Margen por Mes
add_heading('2.1  Ventas y Margen Bruto por Mes', 2)
add_table(
    ['Mes', 'Vtas 2025 $M', 'Vtas 2026 $M', 'Crec. $M', 'Crec. %', 'MB% 2025', 'MB% 2026', 'Var pp', 'U.Neta% 2025', 'U.Neta% 2026'],
    [
        ['Enero',     '$4,449', '$3,718', '-$731',   '-16.4%', '37.1%', '32.1%', '-4.9 pp', '22.9%', '14.2%'],
        ['Febrero',   '$4,493', '$5,074', '+$581',   '+12.9%', '36.5%', '30.3%', '-6.1 pp', '22.2%', '14.8%'],
        ['Marzo',     '$4,428', '$5,266', '+$838',   '+18.9%', '32.9%', '25.4%', '-7.4 pp', '22.0%', '-2.9%'],
        ['Abril',     '$3,735', '$4,983', '+$1,248', '+33.4%', '33.5%', '26.9%', '-6.6 pp', '19.4%', '10.0%'],
        ['ACUMULADO', '$17,105','$19,041', '+$1,936', '+11.3%', '35.1%', '28.4%', '-6.6 pp', '21.7%', '8.5%'],
    ],
    col_widths=[2.2, 1.9, 1.9, 1.7, 1.5, 1.5, 1.5, 1.5, 1.8, 1.8]
)
add_body('Tendencia: Enero ya arranca deteriorado (-4.9 pp). Febrero pierde 6.1 pp. Marzo es el peor con -7.4 pp y utilidad neta negativa. Abril mejora levemente pero sigue 6.6 pp por debajo de 2025. El deterioro no mejora — se estabiliza en un nivel estructuralmente inferior.', color=AZUL_OSC)

# 2.2 Mix
add_heading('2.2  Mix de Ventas: Fabricados vs Comercializados', 2)
add_table(
    ['Categoría', 'Vtas 2025 $M', '% 2025', 'Vtas 2026 $M', '% 2026', 'Crecimiento'],
    [
        ['Fabricados totales',  '$15,639', '91.4%', '$15,484', '81.3%', '-1.0%'],
        ['  — Metalmecánica',   '$13,544', '79.2%', '$13,169', '69.2%', '-2.8%'],
        ['  — Plástico',        '$2,096',  '12.2%', '$2,315',  '12.2%', '+10.5%'],
        ['Comercializados',     '$1,767',  '10.3%', '$3,976',  '20.9%', '+125.0%'],
        ['  — Breakers',        '$1,325',  '7.8%',  '$1,713',  '9.0%',  '+29.3%'],
        ['  — Varillas',        '$271',    '1.6%',  '$2,251',  '11.8%', '+731.7%'],
        ['  — Seti',            '$172',    '1.0%',  '$11',     '0.1%',  '-93.5%'],
        ['Descuentos',          '-$380',   '-2.2%', '-$458',   '-2.4%', '+20.4%'],
        ['TOTAL',               '$17,105', '100%',  '$19,041', '100%',  '+11.3%'],
    ],
    col_widths=[4.5, 2.8, 1.8, 2.8, 1.8, 2.5]
)
add_highlight_box('⚠  PUNTO CRÍTICO: los comercializados de bajo margen pasaron del 10.3% al 20.9% del total. Solo ese cambio de mezcla destruye el margen consolidado matemáticamente.', bg_color=NARANJA)

# 2.3 Mix de productos
add_heading('2.3  Análisis del Mix de Productos', 2)
add_table(
    ['Línea', 'Vtas 2025 $M', 'Vtas 2026 $M', 'Crec. %', 'MB% 2025', 'MB% 2026', 'Var pp', 'Impacto UB $M'],
    [
        ['Breakers (comerc.)',  '$1,325', '$1,713', '+29.3%',  '20.3%', '26.1%', '+5.8 pp',  '+$178'],
        ['Fabricado MT',        '$13,544','$13,169','-2.8%',   '38.9%', '36.1%', '-2.8 pp',  '-$518'],
        ['Fabricado Plástico',  '$2,096', '$2,315', '+10.5%',  '32.8%', '18.4%', '-14.4 pp', '-$262'],
        ['Varillas (comerc.)',  '$271',   '$2,251', '+731.7%', '11.7%', '9.1%',  '-2.7 pp',  '+$173'],
        ['Seti',               '$172',   '$11',    '-93.5%',  '21.8%', '32.7%', 'n.a.',     '-$34'],
    ],
    col_widths=[4.0, 2.3, 2.3, 2.0, 1.8, 1.8, 1.8, 2.2]
)
add_body('Líneas con pérdida bruta:', bold=True, color=ROJO)
add_bullet('CE-125: Ventas $196M, Costos $213M → pérdida -$16M (margen -8.3%)', color=ROJO)
add_bullet('Adaptadores: Ventas $128M, Costos $134M → pérdida -$6M (margen -4.7%)', color=ROJO)
add_body('Efecto mix cuantificado: Si en 2026 se hubieran mantenido los márgenes de 2025 por línea, la utilidad bruta sería $6,501M vs la real de $5,835M. Pérdida por deterioro de márgenes: $666M.', bold=True, color=AZUL_OSC)

# 2.4 Costos
add_heading('2.4  Análisis de Costos de Producción — Acumulado Enero-Abril', 2)
add_table(
    ['Componente', '2025 $M', '% Vtas', '2026 $M', '% Vtas', 'Var $M', 'Var %'],
    [
        ['MOD total',              '$1,817', '10.6%', '$2,085', '11.0%', '+$268', '+14.8%'],
        ['  — Personal temporal',  '$1,231', '7.2%',  '$1,455', '7.6%',  '+$224', '+18.2%'],
        ['MOI total',              '$794',   '4.6%',  '$989',   '5.2%',  '+$195', '+24.5%'],
        ['  — Temporal MOI',       '$177',   '1.0%',  '$293',   '1.5%',  '+$116', '+65.1%'],
        ['CIF total',              '$811',   '4.7%',  '$1,235', '6.5%',  '+$424', '+52.3%'],
        ['  — Arriendos bodegas',  '$81',    '0.5%',  '$335',   '1.8%',  '+$254', '+314.9%'],
        ['  — Mant. metalmecánica','$57',    '0.3%',  '$106',   '0.6%',  '+$49',  '+86.6%'],
        ['  — Herramientas',       '$20',    '0.1%',  '$71',    '0.4%',  '+$51',  '+253.0%'],
        ['  — Transportes CIF',    '$26',    '0.1%',  '$46',    '0.2%',  '+$21',  '+81.6%'],
        ['  — Energía eléctrica',  '$231',   '1.4%',  '$213',   '1.1%',  '-$18',  '-8.0%'],
        ['Servicios Producción',   '$107',   '0.6%',  '$132',   '0.7%',  '+$25',  '+23.8%'],
        ['TOTAL COSTOS PROD',      '$3,529', '20.6%', '$4,442', '23.3%', '+$913', '+25.9%'],
    ],
    col_widths=[4.8, 1.8, 1.6, 1.8, 1.6, 1.7, 1.7]
)
add_body('Escalada mensual de costos de producción:', bold=True)
add_bullet('Enero 2026 vs 2025: $876M vs $858M → +2.1% (controlado)')
add_bullet('Febrero 2026 vs 2025: $1,144M vs $847M → +35.1% (disparo arriendos + MOD)', color=NARANJA)
add_bullet('Marzo 2026 vs 2025: $1,201M vs $974M → +23.3% (mant. inyección $52M, predial, tasa seguridad nueva)', color=NARANJA)
add_bullet('Abril 2026 vs 2025: $1,222M vs $851M → +43.6% (mant. metalmec. $63M, herramientas $40M, dotación)', color=ROJO)

page_break()

# 2.5 Estado de Resultados
add_heading('2.5  Estado de Resultados Completo — Acumulado Enero-Abril', 2)
add_table(
    ['Rubro', '2024 $M', '% Vtas', '2025 $M', '% Vtas', '2026 $M', '% Vtas', 'Var $M', 'Var %'],
    [
        ['Ingresos por ventas',    '$15,501','100%', '$17,105','100%', '$19,041','100%',   '+$1,936', '+11.3%'],
        ['  — Fabricados',         '$14,177','91.5%','$15,639','91.4%','$15,484','81.3%',  '-$155',   '-1.0%'],
        ['  — Comercializados',    '$1,518', '9.8%', '$1,767', '10.3%','$3,976', '20.9%', '+$2,209', '+125.0%'],
        ['CMV',                    '$10,074','65.0%','$11,109','65.0%','$13,625','71.6%',  '+$2,516', '+22.7%'],
        ['  — Costos producción',  '$3,675', '23.7%','$3,762', '22.0%','$4,442', '23.3%', '+$680',   '+18.1%'],
        ['  — MP fabricados',      '$5,187', '33.5%','$5,918', '34.6%','$5,863', '30.8%', '-$55',    '-0.9%'],
        ['  — Costo comercializados','$1,213','7.8%','$1,429', '8.4%', '$3,320', '17.4%', '+$1,891', '+132.4%'],
        ['UTILIDAD BRUTA',         '$5,427', '35.0%','$5,996', '35.1%','$5,415', '28.4%', '-$581',   '-9.7%'],
        ['Gastos Administración',  '$458',   '3.0%', '$586',   '3.4%', '$675',   '3.5%',  '+$89',    '+15.2%'],
        ['Gastos Comerciales',     '$650',   '4.2%', '$786',   '4.6%', '$949',   '5.0%',  '+$163',   '+20.8%'],
        ['Gastos Logística-Salida','$683',   '4.4%', '$739',   '4.3%', '$872',   '4.6%',  '+$133',   '+17.9%'],
        ['UTILIDAD OPERACIONAL',   '$3,635', '23.4%','$3,885', '22.7%','$2,919', '15.3%', '-$966',   '-24.9%'],
        ['Ingresos no operac.',    '$494',   '3.2%', '$445',   '2.6%', '$260',   '1.4%',  '-$185',   '-41.5%'],
        ['  — Financieros',        '$337',   '—',    '$173',   '—',    '$85',    '—',     '-$88',    '-51.0%'],
        ['  — Arrend. bodega rec.',  '$129', '—',    '$117',   '—',    '$155',   '—',     '+$38',    '+32.5%'],
        ['Gastos no operac.',      '$160',   '1.0%', '$368',   '2.2%', '$522',   '2.7%',  '+$154',   '+41.9%'],
        ['  — Intereses',          '$0',     '—',    '$147',   '—',    '$70',    '—',     '-$77',    '—'],
        ['  — Imp. 4x1000',        '$70',    '—',    '$63',    '—',    '$110',   '—',     '+$47',    '+73.8%'],
        ['  — Extraordinarios',    '$3',     '—',    '$1',     '—',    '$185',   '—',     '+$184',   'n.a.'],
        ['UTILIDAD TERCOL',        '$3,968', '25.6%','$3,962', '23.2%','$2,658', '14.0%', '-$1,304', '-32.9%'],
        ['Gastos Familia',         '$216',   '1.4%', '$251',   '1.5%', '$1,032', '5.4%',  '+$781',   '+311.3%'],
        ['  — De personal familia','$137',   '—',    '$164',   '—',    '$968',   '—',     '+$804',   '+490%'],
        ['UTILIDAD NETA',          '$3,753', '24.2%','$3,712', '21.7%','$1,626', '8.5%',  '-$2,086', '-56.2%'],
    ],
    col_widths=[4.2, 1.6, 1.4, 1.6, 1.4, 1.6, 1.4, 1.6, 1.5]
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 3: CAUSAS PRINCIPALES
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 3: CAUSAS PRINCIPALES DEL DETERIORO (ordenadas por impacto)', 1)

causas = [
    ('CAUSA 1', 'Gastos de la familia', '-$781M (-4.1 pp en utilidad neta)',
     '"De personal familia" saltó de $164M a $968M (+490%). Feb: $267M, Mar: $307M, Abr: $267M. Sin este cambio, la utilidad neta sería $2,407M (12.6%) en lugar de $1,626M. Es la causa con mayor impacto individual y la menos vinculada a la operación del negocio.', ROJO),
    ('CAUSA 2', 'Mezcla hacia comercializados de bajo margen', '~-$666M estimados en Margen Bruto',
     'Las varillas crecieron 731.7% (+$1,980M en ventas) con margen del 9.1%. Ese mismo volumen en fabricados metalmecánicos habría generado ~$860M de UB; en varillas generó solo $204M. Diferencia: $656M que el mix destruyó.', ROJO),
    ('CAUSA 3', 'Colapso del margen de Fabricado Plástico', '-$262M en Margen Bruto',
     'El margen pasó de 32.8% a 18.4% (-14.4 pp). La MP plástica subió de $1,361M a $1,890M (+$529M, +38.9%) mientras las ventas solo crecieron +10.5%. Con el margen histórico, el aporte habría sido $759M en lugar de $426M.', NARANJA),
    ('CAUSA 4', 'Deterioro del margen de Fabricado Metalmecánico', '-$518M en Margen Bruto',
     'Margen bajó de 38.9% a 36.1% (-2.8 pp). La magnitud del volumen ($13,169M) hace que el impacto sea el mayor en términos absolutos. Causa: costos de producción crecieron mientras el volumen fabricado bajó (-2.8%) → mala absorción de costos fijos.', NARANJA),
    ('CAUSA 5', 'Explosión de costos de producción', '+$913M adicionales en CMV',
     'Arriendos familiares: +$254M (+315%) | Herramientas: +$51M (+253%) | Mant. metalmec.: +$49M (+86.6%) | MOD temporales: +$224M (+18.2%) | MOI temporales: +$116M (+65.1%) | Servicios producción: +$25M (+23.8%)', NARANJA),
    ('CAUSA 6', 'Gastos operacionales creciendo sin crecimiento en fabricados', '-$385M en U.Op.',
     'Admón +$89M (+15.2%), comerciales +$163M (+20.8%), logística +$133M (+17.9%). Todos crecen más rápido que el negocio core (fabricados -1%). El personal en todas las áreas creció aunque el volumen fabricado no.', RGBColor(0x70, 0x30, 0xA0)),
    ('CAUSA 7', 'Caída de ingresos no operacionales', '-$185M',
     'Rendimientos financieros: de $173M a $85M (-51%). La caja disponible para invertir se redujo, probablemente por mayor capital de trabajo comprometido en inventario de varillas e insumos.', RGBColor(0x70, 0x30, 0xA0)),
]

for num, titulo, impacto, desc, color in causas:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f'{num}: ')
    r1.font.size = Pt(11); r1.font.bold = True; r1.font.color.rgb = color
    r2 = p.add_run(titulo)
    r2.font.size = Pt(11); r2.font.bold = True; r2.font.color.rgb = AZUL_OSC
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(1)
    r3 = p2.add_run(f'Impacto: {impacto}')
    r3.font.size = Pt(10); r3.font.bold = True; r3.font.color.rgb = color
    add_body(desc, indent=1)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 4: RIESGOS
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 4: RIESGOS DEL NEGOCIO', 1)

add_heading('Riesgos Comerciales', 2)
add_bullet('Dependencia de varillas con margen 9.1%: Si las importaciones continúan creciendo, el mix seguirá deteriorando el margen. Cualquier presión de precios convierte este producto en uno con pérdida neta.', color=ROJO)
add_bullet('Caída en ventas metalmecánica fabricada (-2.8%, -$374M): El corazón del negocio pierde momentum. ¿Pérdida de participación de mercado? ¿Competencia con importados?', color=ROJO)
add_bullet('CE-125 con margen -8.3% y Adaptadores con -4.7%: Dos líneas que se fabrican a pérdida bruta (-$22M combinados).', color=ROJO)
add_bullet('Descuentos creciendo al 20.4% (+$78M): Señal de presión comercial o política de precios permisiva.', color=NARANJA)

add_heading('Riesgos Operativos', 2)
add_bullet('Alta dependencia de personal temporal en MOD (69.8% del MOD total, $1,455M): Con 104 temporales en abril, cualquier alza del salario mínimo impacta directamente la estructura.', color=NARANJA)
add_bullet('MOI sobre-estructurado: De 36 a 42 personas (+6 cargos). Costo anualizado: +$195M. El retorno de esta inversión no es visible aún en productividad.', color=NARANJA)
add_bullet('Mantenimientos imprevistos en Plástico: Reforma molde 225A ($43M) y caja oct8 ($9M) en Marzo. Activos con vida útil extendida que se cargan al período.', color=NARANJA)
add_bullet('Mala absorción de costos fijos: La planta absorbe más costo fijo (arriendos $335M, depreciación $155M) con menos volumen fabricado → costo fijo por unidad crece automáticamente.', color=NARANJA)

add_heading('Riesgos Financieros', 2)
add_bullet('Gastos familia $1,032M en 4 meses ($258M/mes promedio): A este ritmo el año cierra con $2,400–$3,000M en gastos familia. Hace inviable cualquier rentabilidad sostenible.', color=ROJO)
add_bullet('$185M en "extraordinarios" no operacionales (2026 vs $1M en 2025): Solo en Abril: $149M. Naturaleza desconocida. Requiere explicación urgente.', color=ROJO)
add_bullet('$408M de intereses en Marzo 2026 (un solo mes): En 2025 el total acumulado fue $147M. ¿Devengo anual? ¿Deuda nueva? La posición financiera necesita revisión urgente.', color=ROJO)
add_bullet('Ingresos financieros cayendo -51% (-$88M): Mayor capital de trabajo comprometido en varillas reduce la caja disponible para inversión.', color=NARANJA)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 5: PLAN DE ACCIÓN
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 5: PLAN DE ACCIÓN', 1)

planes = [
    ('30 DÍAS — Detener la hemorragia', ROJO, [
        ('1. Auditar y controlar los gastos de la familia.',
         'En 2025: $41M/mes promedio. En 2026: $258M/mes promedio. Establecer límite máximo de $80-100M/mes hasta que la utilidad neta recupere el 15%+. Sin esta decisión, todo lo demás es cosmético.'),
        ('2. Establecer precio mínimo de venta para varillas.',
         'Calcular el costo completo: importación + aranceles + fletes + financiación del inventario a tasa vigente + asignación de gastos generales. Con $2,251M en ventas al 9.1% de margen, cualquier costo no asignado lleva a pérdida neta.'),
        ('3. Solicitar costeo real de CE-125 y Adaptadores.',
         'Si el precio sigue por debajo del costo ($22M en pérdida bruta combinada): ajuste de precio inmediato o suspensión de producción.'),
        ('4. Congelar nuevas contrataciones en MOI.',
         'Hasta validar el ROI de las 6 posiciones adicionales (+$195M anualizados) en logística, innovación, mantenimiento y calidad.'),
        ('5. Revisar el contrato de arriendos familiares.',
         'El canon pasó de ~$20M/mes a ~$85M/mes (+$254M en el período). Solicitar contrato y avalúo comercial. Comparar con precios de mercado industrial.'),
    ]),
    ('60 DÍAS — Estabilizar márgenes', NARANJA, [
        ('6. Implementar costeo real por línea de producto mensualmente.',
         'El sistema debe producir el margen bruto por línea el día 5 de cada mes. Sin este dato, no hay toma de decisiones posible.'),
        ('7. Revisión de la estructura de precios del Fabricado Plástico.',
         'Identificar si el costo de insumos plásticos subió (+$529M en MP plástica, +38.9%) por impuesto al plástico, tipo de cambio en resinas, cambio de proveedor, o si hay un problema de eficiencia/scrap en la inyectora.'),
        ('8. Auditoría de los extraordinarios de $185M.',
         'Identificar cada concepto. En 2025 este rubro era prácticamente cero ($1M). Si son provisiones por litigios o pérdidas en activos, registrar adecuadamente.'),
        ('9. Plan de recuperación en metalmecánica.',
         'Identificar por qué las ventas fabricadas cayeron -2.8% (-$374M). Plan con objetivos específicos por vendedor y por cliente.'),
    ]),
    ('90 DÍAS — Reestructurar el modelo', VERDE, [
        ('10. Definir la estrategia de comercializados.',
         'El negocio fabrica al 38.9% de margen y comercializa varillas al 9.1%. Por cada $100 vendidos en varillas, el negocio gana $9 brutos; por $100 en metalmecánica, gana $39. Establecer margen mínimo de admisión para cualquier comercializado.'),
        ('11. Plan de eficiencia en costos de producción.',
         'El costo de producción creció +25.9% (+$913M) con volumen de fabricados plano. Revisar tiempos estándar, índices de scrap, tiempo muerto y productividad por proceso.'),
        ('12. Revisión estructural de gastos familiares.',
         'Comparar los $968M de "de personal familia" con compensaciones ejecutivas de mercado. Definir una política formal compatible con la rentabilidad del negocio.'),
    ]),
]

for titulo, color, items in planes:
    add_highlight_box(titulo, bg_color=color)
    for accion, detalle in items:
        add_body(accion, bold=True)
        add_body(detalle, indent=1)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 6: TABLERO DE INDICADORES
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 6: TABLERO DE INDICADORES RECOMENDADO', 1)
add_table(
    ['Indicador', 'Frecuencia', 'Alerta Roja', 'Alerta Amarilla', 'Meta 2026'],
    [
        ['Ventas totales',             'Semanal',   '<$1,300M/mes', '<$1,500M/mes', '>$5,500M/mes'],
        ['MB% consolidado',            'Mensual',   '<28%',         '<32%',         '≥35%'],
        ['MB% metalmecánica',          'Mensual',   '<33%',         '<36%',         '≥39%'],
        ['MB% plástico',               'Mensual',   '<15%',         '<25%',         '≥30%'],
        ['MB% varillas',               'Mensual',   '<8%',          '<11%',         '≥11%'],
        ['MB% breakers',               'Mensual',   '<20%',         '<23%',         '≥25%'],
        ['Mix comerc./total',          'Mensual',   '>25%',         '>18%',         '≤15%'],
        ['MOD / Ventas',               'Mensual',   '>13%',         '>11.5%',       '≤10.6%'],
        ['MOI / Ventas',               'Mensual',   '>6%',          '>5.5%',        '≤4.6%'],
        ['CIF (sin arriendos) / Vtas', 'Mensual',   '>5%',          '>4.5%',        '≤4.5%'],
        ['Arriendos netos / mes',      'Mensual',   '>$60M',        '>$45M',        '≤$20M'],
        ['Logística / Ventas',         'Mensual',   '>6%',          '>5%',          '≤4.3%'],
        ['Gtos familia / Ventas',      'Mensual',   '>4%',          '>2%',          '≤1.5%'],
        ['U.Operacional %',            'Mensual',   '<12%',         '<18%',         '≥22%'],
        ['U.Neta %',                   'Mensual',   '<6%',          '<14%',         '≥20%'],
        ['# Personal MOI total',       'Mensual',   '>44',          '>42',          '≤38'],
        ['# Personal MOD temporal',    'Semanal',   '>115',         '>110',         '≤105'],
    ],
    col_widths=[5.0, 2.2, 2.8, 2.8, 2.8]
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 7: HALLAZGOS ACCIONABLES
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 7: HALLAZGOS ACCIONABLES — RANKING DE PRIORIDADES', 1)

hallazgos_acc = [
    ('CRÍTICO 1', 'Gastos de la familia', ROJO, [
        ('Qué pasa', '"De personal familia" saltó de $164M a $968M en 4 meses (+$804M, +490%). Feb: $267M, Mar: $307M, Abr: $267M.'),
        ('Evidencia', 'Sin este cambio la utilidad neta sería $2,407M (12.6%) en lugar de $1,626M (8.5%). Impacto directo: -$781M.'),
        ('Hipótesis', 'Nueva política de remuneración a propietarios/familia, o incorporación de nuevos miembros a nómina a costos muy elevados.'),
        ('Acción', 'Auditar cada contrato/acuerdo. Límite máximo: 25% de la utilidad operacional mensual.'),
    ]),
    ('CRÍTICO 2', 'Varillas — volumen masivo con margen mínimo', ROJO, [
        ('Qué pasa', 'De $271M a $2,251M (+731.7%). Margen cayó de 11.7% a 9.1%. En Febrero: solo 7.6%.'),
        ('Evidencia', 'Por cada $100 de varilla vendida, el negocio gana $9 brutos. El mismo esfuerzo en fabricados generaría $39.'),
        ('Hipótesis', 'Crecimiento oportunista en importados sin análisis de rentabilidad integral (costo de capital, fletes, riesgo cambiario).'),
        ('Acción', 'Calcular costo completo vs precio de venta. Establecer precio mínimo. Reducir próxima orden si no es rentable.'),
    ]),
    ('CRÍTICO 3', 'Fabricado Plástico — derrumbe de margen', ROJO, [
        ('Qué pasa', 'Margen de 32.8% cayó a 18.4% (-14.4 pp). MP plástica: $1,361M → $1,890M (+$529M, +38.9%).'),
        ('Evidencia', 'Con el margen de 32.8%, el aporte sería $759M. El real fue $426M. Diferencia: -$333M en solo esta línea.'),
        ('Hipótesis', 'Subida de resinas; impuesto al plástico; cambio de proveedor; scrap elevado; reforma de molde cargada al período ($52M en Marzo).'),
        ('Acción', 'Costeo por referencia plástica. Comparar precio de insumos mes a mes. Identificar las 5 referencias con mayor deterioro.'),
    ]),
    ('CRÍTICO 4', 'CE-125 y Adaptadores — pérdida bruta', ROJO, [
        ('Qué pasa', 'CE-125 margen -8.3% (-$16M). Adaptadores margen -4.7% (-$6M). La empresa paga $22M para fabricar estos productos.'),
        ('Evidencia', 'CE-125: costo $213M > ventas $196M. Adaptadores: costo $134M > ventas $128M.'),
        ('Hipótesis', 'Precio de venta desactualizado respecto al costo actual de inyección + metalización + ensamble.'),
        ('Acción', 'Costo unitario estándar vs precio de lista. Si persiste la pérdida: ajuste de precio esta semana o discontinuación.'),
    ]),
    ('IMPORTANTE 5', 'Arriendos familiares', NARANJA, [
        ('Qué pasa', 'Canon mensual pasó de ~$20M a ~$85M/mes. Total período: $81M (2025) → $335M (2026). Diferencia: +$254M.'),
        ('Evidencia', 'Neto (descontando ingreso de arriendo recibido $155M): +$216M adicionales que no existían en 2025.'),
        ('Hipótesis', 'A partir de agosto 2025 se incorporaron bodegas de la familia al contrato a precios potencialmente sobre mercado.'),
        ('Acción', 'Obtener avalúo comercial de los inmuebles y comparar con canon actual. Negociar ajuste si está fuera de mercado.'),
    ]),
    ('IMPORTANTE 6', 'Mantenimientos y herramientas', NARANJA, [
        ('Qué pasa', 'Mant. metalmecánica +$49M (+86.6%), herramientas +$51M (+253%). Solo en Abril: mant. $63M + herramientas $40M.'),
        ('Evidencia', 'Reforma molde BSC225T $36M, repuesto punzonadora $5M, servomotor + herramientas Amada.'),
        ('Hipótesis', 'Maquinaria envejecida o período de rodaje costoso del equipo Amada (consumibles de alta tecnología).'),
        ('Acción', 'Presupuesto mensual de mantenimiento y herramientas. Evaluar si la inversión en Amada genera productividad que justifique el costo.'),
    ]),
    ('SEGUIMIENTO 7', 'Ingresos financieros cayendo', RGBColor(0x70, 0x30, 0xA0), [
        ('Qué pasa', 'De $173M en 2025 a $85M en 2026 (-$88M, -51%).'),
        ('Evidencia', 'Señal de menor caja disponible para inversión.'),
        ('Hipótesis', 'Mayor capital de trabajo comprometido en inventario de varillas reduce la caja disponible.'),
        ('Acción', 'Revisar el ciclo de conversión de caja, especialmente días de inventario en varillas.'),
    ]),
]

for prioridad, titulo, color, items in hallazgos_acc:
    add_heading(f'{prioridad}: {titulo}', 3)
    for etiqueta, texto in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f'{etiqueta}: ')
        r1.font.size = Pt(10); r1.font.bold = True; r1.font.color.rgb = color
        r2 = p.add_run(texto)
        r2.font.size = Pt(10)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PARTE 8: RESPUESTA TIPO DUEÑO/CFO
# ══════════════════════════════════════════════════════════════════════════════
add_heading('PARTE 8: RESPUESTA TIPO DUEÑO / CFO', 1)

add_heading('5 Decisiones para Esta Semana', 2)
decisiones = [
    ('DECISIÓN 1 — HOY', 'Auditar y controlar los gastos de la familia.',
     'En 2025: $41M/mes promedio. En 2026: $258M/mes promedio. Convocar reunión con el dueño/familia, presentar los números crudos, y establecer un límite fijo de $80-100M/mes hasta tanto la utilidad neta no recupere el 15%+ de ventas. Sin esta decisión, todo lo demás es cosmético.'),
    ('DECISIÓN 2 — ESTA SEMANA', 'Precio mínimo de venta para varillas.',
     'Calcular el costo completo: precio de importación + aranceles + fletes + financiación del inventario a tasa vigente + asignación de gastos generales. Con $2,251M en ventas al 9.1% de margen ($204M de UB), cualquier gasto adicional no asignado lleva a pérdida neta. Si el mercado no permite ese mínimo, reducir la próxima orden de importación a la mitad.'),
    ('DECISIÓN 3 — ESTA SEMANA', 'Lista de precios de CE-125 y Adaptadores.',
     'Solicitar el costo unitario estándar actualizado y compararlo con el precio de lista. Si el precio sigue por debajo del costo, comunicar al equipo comercial el nuevo precio mínimo esta misma semana. No aceptar más pedidos al precio actual.'),
    ('DECISIÓN 4 — ESTA SEMANA', 'Revisar el contrato de arriendo de bodegas.',
     'El canon mensual subió de ~$20M a ~$85M/mes (+4.2x en un año). Solicitar el contrato físico, el avalúo comercial, y comparar con 3 cotizaciones de bodegas similares en el mismo sector. Si está fuera de mercado, negociar con la familia una reducción o congelación por 12 meses.'),
    ('DECISIÓN 5 — ESTA SEMANA', 'Congelar contrataciones y revisar MOI.',
     'El personal indirecto subió de 36 a 42 personas (+$195M anualizados). Antes de aprobar cualquier nueva vinculación, cada una de las 6 posiciones adicionales debe tener métricas de desempeño y un análisis de retorno.'),
]
for label, accion, detalle in decisiones:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(2)
    r1 = p.add_run(label + ': ')
    r1.font.size = Pt(10); r1.font.bold = True; r1.font.color.rgb = ROJO
    r2 = p.add_run(accion)
    r2.font.size = Pt(10); r2.font.bold = True; r2.font.color.rgb = AZUL_OSC
    add_body(detalle, indent=1)

add_heading('Indicadores a Monitorear Semanalmente', 2)
add_table(
    ['Indicador', 'Frecuencia', 'Responsable'],
    [
        ['Ventas por línea (fabricados vs comercializados)',       'Semanal viernes',  'Gerente Comercial'],
        ['Mix comercializados / total ventas (hoy: 20.9%)',        'Semanal',          'Gerente Comercial'],
        ['Margen bruto % estimado por línea',                     'Quincenal',        'Controller'],
        ['Costos de producción vs presupuesto (hoy: $1,110M/mes)','Mensual día 5',    'Jefe Prod. / Controller'],
        ['# personas MOD y MOI (hoy: 104 + 42)',                  'Semanal lunes',    'RRHH'],
        ['Gastos familiares del mes (hoy: $258M/mes promedio)',    'Mensual',          'Gerente Financiero'],
        ['Precio varillas vs costo real (incluyendo financiación)','Semanal',          'Gerente Compras'],
        ['Días de inventario + días de cartera',                   'Semanal',          'Tesorero'],
    ],
    col_widths=[8.0, 3.0, 4.6]
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  CONCLUSIÓN FINAL
# ══════════════════════════════════════════════════════════════════════════════
add_heading('CONCLUSIÓN FINAL', 1)

add_highlight_box(
    'El margen está cayendo principalmente por X, Y y Z:',
    bg_color=AZUL_OSC
)

add_body('X — Cambio radical de mix hacia varillas comercializadas al 9.1% de margen', bold=True, color=ROJO)
add_body('Las varillas pasaron de $271M (1.6% de ventas) a $2,251M (11.8% de ventas). Solo este efecto mezcla destruye 3-4 puntos porcentuales del margen bruto consolidado: se vende mucho más, pero cada peso adicional vendido en varillas aporta 4 veces menos utilidad que un peso en metalmecánica fabricada.', indent=1)

add_body('Y — Explosión de costos de producción (+25.9%, +$913M) sobre un volumen de fabricados plano o en caída (-2.8%, -$374M)', bold=True, color=ROJO)
add_body('Liderado por arriendos familiares que se multiplicaron por 4 (+$254M), herramientas y mantenimientos que casi se triplicaron (+$100M), y contrataciones en MOI que no muestran aún retorno en productividad (+$195M anualizados).', indent=1)

add_body('Z — Gastos de la familia que saltaron +311% (+$781M): de $251M a $1,032M en los mismos 4 meses', bold=True, color=ROJO)
add_body('El promedio mensual pasó de $63M a $258M. El negocio opera razonablemente bien a nivel de utilidad operacional ($2,919M, 15.3%), pero transfiere $1,032M fuera de la operación vía este renglón, dejando solo $1,626M (8.5%) como utilidad neta — menos de la mitad de lo que quedaba en 2025.', indent=1)

doc.add_paragraph()
add_highlight_box(
    'La situación es seria pero recuperable. El negocio core (metalmecánica fabricada) todavía opera al 36.1% de margen bruto, la capacidad instalada existe, y el crecimiento de ventas demuestra que hay demanda. Las tres causas identificadas tienen solución ejecutiva directa — no requieren una transformación de largo plazo, requieren decisiones esta semana.',
    bg_color=VERDE
)

# ══════════════════════════════════════════════════════════════════════════════
#  5 PREGUNTAS CRÍTICAS
# ══════════════════════════════════════════════════════════════════════════════
add_heading('5 PREGUNTAS CRÍTICAS PARA INVESTIGAR INTERNAMENTE', 1)

preguntas = [
    '1. ¿Por qué "de personal familia" subió de $164M a $968M en los mismos 4 meses (+$804M)? ¿Se incorporaron nuevos miembros a nómina, se aumentaron salarios de los existentes, o se están pagando dividendos/anticipos/reembolsos a través de nómina? ¿Existe una política aprobada por junta directiva para esto?',
    '2. ¿Cuál es el plan de compra de varillas para los próximos 6 meses y cuál es el precio mínimo de venta aprobado? ¿Quién tomó la decisión de crecer de $271M a $2,251M en varillas (+731.7%) sin calcular el impacto en el margen consolidado? ¿Existe un análisis de rentabilidad integral incluyendo el costo financiero del inventario?',
    '3. ¿Cuál es la causa exacta del derrumbe de margen en Fabricado Plástico (32.8% → 18.4%)? La MP plástica subió $529M (+38.9%) pero las ventas solo crecieron +$219M (+10.5%). ¿Es precio de resinas, impuesto al plástico, scrap elevado, cambio de proveedor, o carga de reforma de moldes al período?',
    '4. ¿Qué son los $185M en "extraordinarios" de gastos no operacionales y los $408M en intereses de Marzo 2026? En 2025 el total de intereses acumulados en 4 meses fue $147M. ¿Son ajustes contables puntuales? ¿La empresa tiene deuda financiera nueva?',
    '5. ¿Cuál es la justificación del canon de arriendo de bodegas familiares a $85M/mes? ¿Se realizó un avalúo comercial? ¿El valor está dentro del rango de mercado para bodegas industriales en esa zona? ¿El ingreso de $155M por arrendamiento que aparece como no operacional corresponde a las mismas bodegas u otras?',
]

for preg in preguntas:
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    c0 = table.rows[0].cells[0]
    c1 = table.rows[0].cells[1]
    c0.text = '?'
    c0.paragraphs[0].runs[0].font.size = Pt(16)
    c0.paragraphs[0].runs[0].font.bold = True
    c0.paragraphs[0].runs[0].font.color.rgb = BLANCO
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_bg(c0, AZUL_MED)
    c0.width = Cm(0.8)
    c1.text = preg
    c1.paragraphs[0].runs[0].font.size = Pt(10)
    set_cell_bg(c1, GRIS_CLARO)
    set_cell_borders(c0, 'FFFFFF', 2)
    set_cell_borders(c1, 'CCCCCC', 2)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ══════════════════════════════════════════════════════════════════════════════
#  PIE DE PÁGINA
# ══════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Análisis construido exclusivamente sobre los datos de los archivos proporcionados. Todas las cifras en millones de pesos colombianos (COP $M).')
run.font.size = Pt(8)
run.font.italic = True
run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ══════════════════════════════════════════════════════════════════════════════
#  GUARDAR
# ══════════════════════════════════════════════════════════════════════════════
output_path = '/home/user/analisis-financiero/Diagnostico_Financiero_2026.docx'
doc.save(output_path)
print(f'Archivo generado: {output_path}')
