#!/usr/bin/env python3
"""
Generador de Informe Mensual DOCX — DC Engineering Group
Genera el informe completo con todas las secciones automáticamente.
"""

import io, os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), kwargs.get(edge, 'single'))
        tag.set(qn('w:sz'), '4')
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), '000000')
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1'] if level == 1 else doc.styles['Heading 2']
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(12)
    else:
        run.font.size = Pt(11)
    return p

def add_para(doc, text, bold=False, italic=False, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.alignment = alignment
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

def add_table_header_row(table, headers, bg_color='1F497D'):
    row = table.rows[0] if table.rows else table.add_row()
    for i, header in enumerate(headers):
        if i < len(row.cells):
            cell = row.cells[i]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(header)
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = Pt(9)
            set_cell_bg(cell, bg_color)

def generar_informe_mensual(data):
    """
    Genera el DOCX del informe mensual completo.
    data: dict con todos los datos del informe
    """
    doc = Document()
    
    # Márgenes
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    # ── PORTADA ──────────────────────────────────────────────────────────────
    # Agencia
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(data.get('agencia', 'Municipio Autónomo de Caguas'))
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    doc.add_paragraph()

    # Nombre del proyecto
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(data.get('nombre_proyecto', '').upper())
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()

    # Título del informe
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(data.get('titulo_informe', f"Informe Mensual #{data.get('numero_informe', 1)}"))
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()

    # DC Engineering info
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('DC ENGINEERING GROUP, PSC')
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('First Federal Savings Building Suite #510, Ponce de León Ave. #1519, SJ, PR, 00909').font.size = Pt(9)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('Tel: 787-477-1789 | dcordova@dc-eng.com').font.size = Pt(9)

    doc.add_page_break()

    # ── TABLA DE CONTENIDO ────────────────────────────────────────────────────
    add_para(doc, 'Table of Contents', bold=True, size=12)
    toc_items = [
        ('I.', 'Descripción del Proyecto'),
        ('II.', 'Información General del Proyecto'),
        ('III.', 'Resumen de Trabajos Realizados'),
        ('IV.', 'Análisis de Costos y Tiempo'),
        ('V.', 'Fianzas y Seguros'),
        ('VI.', 'Resumen de Facturación'),
        ('VII.', 'Request for Information'),
        ('VIII.', 'Permisos y Endosos'),
        ('IX.', 'Submittal'),
        ('X.', 'Anejo A: Comunicación Escrita Enviada/Recibida'),
        ('XI.', 'Anejo B: Submittals'),
        ('XII.', 'Anejo C: Informes Diarios de Inspección'),
    ]
    for num, titulo in toc_items:
        p = doc.add_paragraph()
        p.add_run(num).bold = True
        p.add_run(f'\t{titulo}').font.size = Pt(10)

    doc.add_page_break()

    # ── I. DESCRIPCIÓN DEL PROYECTO ───────────────────────────────────────────
    add_heading(doc, 'Descripción del Proyecto')
    nombre_proy = data.get('nombre_proyecto', '')
    add_para(doc, f'El proyecto propuesto consiste en la {nombre_proy}')

    # ── II. INFORMACIÓN GENERAL DEL PROYECTO ─────────────────────────────────
    add_heading(doc, 'Información General del Proyecto')
    add_para(doc, 'Partes Involucradas en el Proyecto de Referencia:')

    t = doc.add_table(rows=1, cols=2)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.columns[0].width = Inches(1.8)
    t.columns[1].width = Inches(4.7)
    add_table_header_row(t, ['Participantes', 'Descripción'])

    partes = [
        ('Dueño', f"{data.get('agencia', 'MUNICIPIO AUTONOMO DE CAGUAS')}\nAgrim. Yamil Sanabria Roman\nDirector Interino Departamento de Obras Públicas Municipio Autónomo de Caguas\nEmail: Yamil.Sanabria@caguas.gov.pr"),
        ('Inspección', f"DC ENGINEERING GROUP, PSC\nFirst Federal Savings Building Suite #510, Ponce de León Ave. #1519, SJ, PR, 00909\nProject Inspector: {data.get('inspector', 'Ing. Francisco Cordova, PE')} – Email: fcordova@dc-eng.com\nTel: 787-318-0746"),
    ]
    
    diseñador = data.get('disenador', '')
    if diseñador:
        partes.append(('Diseño', diseñador))
    
    contratista = data.get('contratista', '')
    contratista_contacto = data.get('contratista_contacto', '')
    partes.append(('Contratista General', f"{contratista}\n{contratista_contacto}" if contratista_contacto else contratista))

    for rol, descripcion in partes:
        row = t.add_row()
        row.cells[0].text = rol
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].text = descripcion

    # ── III. RESUMEN DE TRABAJOS REALIZADOS ───────────────────────────────────
    add_heading(doc, 'Resumen de Trabajos Realizados')

    add_para(doc, 'Comunicación Escrita', bold=True)
    comunicaciones = data.get('comunicaciones', [])
    if comunicaciones:
        add_para(doc, 'Durante el periodo que comprende este informe se generaron las siguientes comunicaciones escritas:')
        for com in comunicaciones:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(com).font.size = Pt(10)
    else:
        add_para(doc, 'Durante el periodo que comprende este informe no se recibieron y/o enviaron comunicaciones escritas.')
    
    add_para(doc, 'Para más detalles ver Anejo A – Comunicaciones escritas', italic=True)

    add_para(doc, 'Tareas De Campo:', bold=True)
    resumen_act = data.get('resumen_actividades', 'Para más detalles ver Anejo C – Informes Semanales y Diarios')
    # Split por saltos de línea o por oraciones largas
    for linea in resumen_act.split('\n'):
        linea = linea.strip()
        if linea:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(linea).font.size = Pt(10)
    
    add_para(doc, 'Para más detalles ver Anejo C – Informes Semanales y Diarios', italic=True)

    # ── IV. ANÁLISIS DE COSTOS Y TIEMPO ──────────────────────────────────────
    add_heading(doc, 'Análisis de Costos y Tiempo')
    add_para(doc, f'A continuación, se detalla el análisis de costos y tiempo del contratista {contratista}.')
    add_para(doc, contratista, bold=True)
    add_para(doc, f'Número de Contrato: {data.get("numero_contrato", "N/A")}')

    t = doc.add_table(rows=3, cols=2)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.columns[0].width = Inches(1.5)
    t.columns[1].width = Inches(5.0)

    # COST row
    t.cell(0, 0).text = 'COST'
    t.cell(0, 0).paragraphs[0].runs[0].bold = True
    t.cell(0, 1).text = f"PO {data.get('po_number', 'N/A')}"
    # Agregar el monto en la misma celda
    p2 = t.cell(0, 1).add_paragraph(data.get('costo_contrato', 'N/A'))
    p2.runs[0].bold = True

    # TIME row
    t.cell(1, 0).text = 'TIME'
    t.cell(1, 0).paragraphs[0].runs[0].bold = True
    t.cell(1, 1).text = 'Project Duration:'
    t.cell(2, 0).text = ''
    dias = data.get('dias_contrato', 'N/A')
    t.cell(2, 1).text = f'Contrato por {dias}' if 'mes' not in str(dias).lower() and 'día' not in str(dias).lower() else str(dias)

    # ── V. FIANZAS Y SEGUROS ──────────────────────────────────────────────────
    add_heading(doc, 'Fianzas y Seguros')
    add_para(doc, f'Las tablas a continuación resumen los números y fechas de vigencia de los seguros del contratista {contratista}.')
    add_para(doc, contratista, bold=True)

    seguros = data.get('seguros', [])
    cols_seg = ['Tipo De Póliza', 'Número De Póliza', 'Fecha De Emisión', 'Fecha De Vencimiento', 'Límite']
    t = doc.add_table(rows=1, cols=len(cols_seg))
    t.style = 'Table Grid'
    add_table_header_row(t, cols_seg)

    if seguros:
        for seg in seguros:
            row = t.add_row()
            row.cells[0].text = seg.get('tipo', '-')
            row.cells[1].text = seg.get('numero', '-')
            row.cells[2].text = seg.get('fecha_emision', '-')
            row.cells[3].text = seg.get('fecha_vencimiento', '-')
            row.cells[4].text = seg.get('limite', '-')
    else:
        row = t.add_row()
        for i in range(5): row.cells[i].text = '-'
        add_para(doc, '(Los datos de seguros se encuentran en el Anejo A del informe anterior)', italic=True)

    # ── VI. RESUMEN DE FACTURACIÓN ────────────────────────────────────────────
    add_heading(doc, 'Resumen de Facturación')
    add_para(doc, f'Las tablas a continuación resumen de las facturas del contratista {contratista}.')
    add_para(doc, contratista, bold=True)

    cols_fac = ['Numero', 'Contratista', 'Cantidad', 'Porciento', 'Fecha Sometida A Administración', 'Status']
    t = doc.add_table(rows=1, cols=len(cols_fac))
    t.style = 'Table Grid'
    add_table_header_row(t, cols_fac)

    certificaciones = data.get('certificaciones', [])
    if certificaciones:
        for cert in certificaciones:
            row = t.add_row()
            row.cells[0].text = str(cert.get('numero', '-'))
            row.cells[1].text = cert.get('contratista', contratista)
            row.cells[2].text = cert.get('cantidad', '-')
            row.cells[3].text = cert.get('porciento', '-')
            row.cells[4].text = cert.get('fecha', '-')
            row.cells[5].text = cert.get('status', 'Pendiente')
    else:
        row = t.add_row()
        vals = ['1', contratista, '-', '-', '-', '-']
        for i, v in enumerate(vals): row.cells[i].text = v

    # ── VII. REQUEST FOR INFORMATION ──────────────────────────────────────────
    add_heading(doc, 'Request for Information')
    rfis = data.get('rfis', [])
    
    if rfis:
        add_para(doc, f'Durante el periodo que comprende este informe, se recibieron los siguientes RFIs.')
    else:
        add_para(doc, 'Durante el periodo que comprende este informe, se recibieron los siguientes RFIs.')
    
    add_para(doc, contratista, bold=True)
    cols_rfi = ['No.', 'RFI No.', 'Description', 'Date Sent']
    t = doc.add_table(rows=1, cols=len(cols_rfi))
    t.style = 'Table Grid'
    add_table_header_row(t, cols_rfi)
    
    if rfis:
        for rfi in rfis:
            row = t.add_row()
            row.cells[0].text = str(rfi.get('no', '-'))
            row.cells[1].text = str(rfi.get('rfiNo', '-'))
            row.cells[2].text = rfi.get('descripcion', '-')
            row.cells[3].text = rfi.get('fecha', '-')
    else:
        row = t.add_row()
        for i, v in enumerate(['1', '-', '-', '-']): row.cells[i].text = v

    # ── VIII. PERMISOS Y ENDOSOS ──────────────────────────────────────────────
    add_heading(doc, 'Permisos y Endosos')
    add_para(doc, 'La tabla a continuación resume permisos expedidos para el proyecto y ofrece una descripción de estos. Durante el periodo que comprende este informe no se generaron permisos.')

    cols_per = ['Descripción', 'Agencia', 'ESTATUS', 'Número de Caso']
    t = doc.add_table(rows=1, cols=len(cols_per))
    t.style = 'Table Grid'
    add_table_header_row(t, cols_per)
    row = t.add_row()
    for i in range(4): row.cells[i].text = '-'

    # ── IX. SUBMITTAL ─────────────────────────────────────────────────────────
    add_heading(doc, 'Submittal')
    add_para(doc, 'La tabla a continuación contiene todos los submittals del proyecto de referencia.')
    add_para(doc, contratista, bold=True)

    cols_sub = ['No.', 'Submittal No.', 'Title', 'Status', 'Received from Contractor', 'Sent to Designer', 'Returned from Designer', 'Sent to Contractor']
    t = doc.add_table(rows=1, cols=len(cols_sub))
    t.style = 'Table Grid'
    # Ajustar anchos
    widths = [0.35, 0.6, 2.0, 0.7, 0.8, 0.7, 0.85, 0.7]
    for i, w in enumerate(widths):
        t.columns[i].width = Inches(w)
    add_table_header_row(t, cols_sub)

    submittals = data.get('submittals_acumulativos', [])
    if submittals:
        for idx, s in enumerate(submittals, 1):
            row = t.add_row()
            row.cells[0].text = str(idx)
            row.cells[1].text = str(s.get('no', idx))
            row.cells[2].text = s.get('titulo', '-')
            row.cells[3].text = s.get('status', '-')
            row.cells[4].text = s.get('fechaRecibido', '-')
            row.cells[5].text = s.get('fechaDiseñador', '-')
            row.cells[6].text = s.get('fechaRetornado', '-')
            row.cells[7].text = s.get('fechaContratista', '-')
    else:
        row = t.add_row()
        for i in range(8): row.cells[i].text = '-'

    # ── X. ANEJO A ────────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, 'Anejo A: Comunicación Escrita Enviada/Recibida')
    
    if comunicaciones:
        for com in comunicaciones:
            p = doc.add_paragraph(style='List Number')
            p.add_run(com).font.size = Pt(10)
    else:
        add_para(doc, 'Durante el periodo que comprende este informe no se recibieron Comunicaciones escritas.')

    # ── XI. ANEJO B ───────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, 'Anejo B: Submittals')
    
    submittals_periodo = data.get('submittals_periodo', [])
    if submittals_periodo:
        for s in submittals_periodo:
            p = doc.add_paragraph(style='List Number')
            p.add_run(s).font.size = Pt(10)
    else:
        add_para(doc, 'Durante el periodo que comprende este informe no se recibieron Submittals.')

    # ── XII. ANEJO C ──────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, 'Anejo C: Informes Diarios de Inspección')
    add_para(doc, f'Se adjuntan los informes diarios de inspección correspondientes al período {data.get("periodo_texto", "")}.')
    add_para(doc, '(Adjuntar PDF de informes diarios del período)', italic=True)

    # Guardar en buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
