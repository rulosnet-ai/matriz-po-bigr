"""
Matriz de Producción - Generador automático desde PDF de Orden de Compra
Cliente objetivo: Big R (formato estándar de PO)

Cómo correr:
    streamlit run app.py
"""

import re
import io
from collections import defaultdict

import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ----------------------------------------------------------------------------
# EXTRACCIÓN DE DATOS DEL PDF
# ----------------------------------------------------------------------------

# Patrón de una línea de detalle, por ejemplo:
# 1 1 1 2129550 M 5 PCKT JEAN MSH 29X30 BR39-AFJNG-MSH 18.60 EA 18.60
LINE_PATTERN = re.compile(
    r"^(?P<line_no>\d+)\s+"
    r"(?P<store>\d+)\s+"
    r"(?P<qty>\d+)\s+"
    r"(?P<sku>\d{5,})\s+"
    r"(?P<desc>.+?)\s+"
    r"(?P<size>\d{2}X\d{2})\s+"
    r"(?P<mfg>[A-Z0-9\-]+)\s+"
    r"(?P<unit_cost>\d+\.\d{2})\s+"
    r"(?P<uom>[A-Z]{2})\s+"
    r"(?P<ext_cost>\d+\.\d{2})\s*$"
)

PO_NUMBER_PATTERN = re.compile(r"P\s*\.?\s*O\s*\.?\s*#\s*:?\s*([\d\s]+)")
CLIENT_NAME_PATTERN = re.compile(r"^(.+?)\s*\nPage:", re.MULTILINE)


def extract_po_metadata(full_text: str) -> dict:
    """Extrae metadatos generales de la orden (PO#, cliente, fechas, total)."""
    meta = {}

    po_match = PO_NUMBER_PATTERN.search(full_text)
    if po_match:
        meta["po_number"] = po_match.group(1).replace(" ", "").strip()

    lines = full_text.splitlines()
    for line in lines:
        stripped = line.strip()
        if "Page:" in stripped:
            meta["client_name"] = stripped.split("Page:")[0].strip()
            break

    date_match = re.search(r"Order Date:\s*([\d/\s]+)", full_text)
    if date_match:
        meta["order_date"] = date_match.group(1).strip()

    total_units_match = re.search(r"TOTAL UNITS\s+(\d+)", full_text)
    if total_units_match:
        meta["total_units_pdf"] = int(total_units_match.group(1))

    total_cost_match = re.search(r"TOTAL P\.O\.\s+([\d,]+\.\d{2})", full_text)
    if total_cost_match:
        meta["total_cost_pdf"] = float(total_cost_match.group(1).replace(",", ""))

    return meta


def extract_lines_from_pdf(file) -> tuple[list[dict], dict, list[str]]:
    """
    Lee el PDF y devuelve:
      - lista de líneas de detalle parseadas
      - metadatos generales de la orden
      - lista de advertencias (líneas que no se pudieron interpretar)
    """
    rows = []
    warnings = []
    full_text_pages = []

    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            full_text_pages.append(text)

            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line or not line[0].isdigit():
                    continue

                match = LINE_PATTERN.match(line)
                if match:
                    d = match.groupdict()
                    cintura, largo = d["size"].split("X")
                    rows.append({
                        "page": page_num,
                        "line_no": int(d["line_no"]),
                        "store": int(d["store"]),
                        "qty": int(d["qty"]),
                        "sku": d["sku"],
                        "desc": d["desc"].strip(),
                        "size": d["size"],
                        "cintura": int(cintura),
                        "largo": int(largo),
                        "mfg": d["mfg"],
                        "unit_cost": float(d["unit_cost"]),
                        "ext_cost": float(d["ext_cost"]),
                    })
                else:
                    # Solo marcar como advertencia si parece una línea de detalle
                    # (empieza con número y tiene varios tokens) pero no calzó el patrón
                    tokens = line.split()
                    if len(tokens) >= 6 and tokens[0].isdigit():
                        warnings.append(f"Página {page_num}: no se pudo interpretar -> {line}")

    full_text = "\n".join(full_text_pages)
    meta = extract_po_metadata(full_text)
    return rows, meta, warnings


# ----------------------------------------------------------------------------
# AGRUPACIÓN Y VALIDACIÓN
# ----------------------------------------------------------------------------

def detect_description_mfg_mismatches(rows: list[dict]) -> list[dict]:
    """
    Detecta líneas donde la descripción (estilo/lavado en texto) no concuerda
    con el grupo dominante de descripciones para ese código MFG.
    Devuelve una lista de inconsistencias para mostrarle al usuario.
    """
    mfg_to_descs = defaultdict(lambda: defaultdict(int))
    for r in rows:
        mfg_to_descs[r["mfg"]][r["desc"]] += 1

    # Descripción dominante por MFG
    dominant_desc = {
        mfg: max(descs.items(), key=lambda kv: kv[1])[0]
        for mfg, descs in mfg_to_descs.items()
    }

    mismatches = []
    for r in rows:
        if r["desc"] != dominant_desc[r["mfg"]]:
            mismatches.append({
                "line_no": r["line_no"],
                "sku": r["sku"],
                "size": r["size"],
                "desc_en_linea": r["desc"],
                "mfg": r["mfg"],
                "desc_esperada_por_mfg": dominant_desc[r["mfg"]],
            })
    return mismatches


def group_by_mfg(rows: list[dict]) -> dict:
    """Agrupa las líneas por código MFG (estilo + lavado de fabricación)."""
    groups = defaultdict(lambda: {"rows": [], "descs": defaultdict(int)})
    for r in rows:
        groups[r["mfg"]]["rows"].append(r)
        groups[r["mfg"]]["descs"][r["desc"]] += 1
    return groups


def build_size_matrix(rows: list[dict]) -> tuple[list[int], list[int], dict]:
    """Construye la matriz cintura x largo -> cantidad para un grupo de líneas."""
    cinturas = sorted(set(r["cintura"] for r in rows))
    largos = sorted(set(r["largo"] for r in rows))
    qty_map = defaultdict(int)
    for r in rows:
        qty_map[(r["cintura"], r["largo"])] += r["qty"]
    return cinturas, largos, qty_map


# ----------------------------------------------------------------------------
# GENERACIÓN DEL EXCEL
# ----------------------------------------------------------------------------

def build_excel(groups: dict, meta: dict, total_qty: int) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"

    title_font = Font(name="Arial", size=14, bold=True)
    subtitle_font = Font(name="Arial", size=10, italic=True)
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2F5496")
    mfg_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    mfg_fill = PatternFill("solid", start_color="595959")
    total_font = Font(name="Arial", size=10, bold=True)
    total_fill = PatternFill("solid", start_color="D9D9D9")
    normal_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws["A1"] = "MATRIZ DE PRODUCCIÓN POR ESTILO/LAVADO - TALLAS (CINTURA x LARGO)"
    ws["A1"].font = title_font

    po = meta.get("po_number", "N/D")
    client = meta.get("client_name", "N/D")
    order_date = meta.get("order_date", "N/D")
    ws["A2"] = f"PO #: {po}  |  Cliente: {client}  |  Fecha orden: {order_date}  |  Total piezas: {total_qty}"
    ws["A2"].font = subtitle_font

    row = 4
    ws.cell(row=row, column=1, value="Código MFG").font = header_font
    ws.cell(row=row, column=2, value="Descripción").font = header_font
    ws.cell(row=row, column=3, value="Total piezas").font = header_font
    for c in (1, 2, 3):
        ws.cell(row=row, column=c).fill = header_fill
    row += 1

    for mfg, g in groups.items():
        total = sum(r["qty"] for r in g["rows"])
        desc = max(g["descs"].items(), key=lambda kv: kv[1])[0]
        ws.cell(row=row, column=1, value=mfg).font = normal_font
        ws.cell(row=row, column=2, value=desc).font = normal_font
        ws.cell(row=row, column=3, value=total).font = normal_font
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = total_font
    ws.cell(row=row, column=3, value=total_qty).font = total_font

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 14

    for mfg, g in groups.items():
        sheet_name = re.sub(r"[^A-Za-z0-9\-]", "", mfg)[:31] or mfg[:31]
        ms = wb.create_sheet(sheet_name)
        desc = max(g["descs"].items(), key=lambda kv: kv[1])[0]
        total = sum(r["qty"] for r in g["rows"])

        ms["A1"] = desc
        ms["A1"].font = title_font
        ms["A2"] = f"Código MFG: {mfg}    |    Total piezas: {total}"
        ms["A2"].font = subtitle_font

        cinturas, largos, qty_map = build_size_matrix(g["rows"])

        start_row = 4
        ms.cell(row=start_row, column=1, value="LARGO \\ CINTURA").font = header_font
        ms.cell(row=start_row, column=1).fill = header_fill
        ms.cell(row=start_row, column=1).border = border
        ms.cell(row=start_row, column=1).alignment = center

        for j, cintura in enumerate(cinturas, start=2):
            cell = ms.cell(row=start_row, column=j, value=cintura)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center

        total_col = len(cinturas) + 2
        cell = ms.cell(row=start_row, column=total_col, value="TOTAL")
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center

        for i, largo in enumerate(largos, start=1):
            r = start_row + i
            cell = ms.cell(row=r, column=1, value=largo)
            cell.font = mfg_font
            cell.fill = mfg_fill
            cell.border = border
            cell.alignment = center

            for j, cintura in enumerate(cinturas, start=2):
                qty = qty_map.get((cintura, largo), 0)
                c = ms.cell(row=r, column=j)
                c.border = border
                c.alignment = center
                c.font = normal_font
                c.value = qty if qty else None

            col_start = get_column_letter(2)
            col_end = get_column_letter(total_col - 1)
            c = ms.cell(row=r, column=total_col, value=f"=SUM({col_start}{r}:{col_end}{r})")
            c.font = total_font
            c.fill = total_fill
            c.border = border
            c.alignment = center

        total_row = start_row + len(largos) + 1
        cell = ms.cell(row=total_row, column=1, value="TOTAL")
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = center

        for j, cintura in enumerate(cinturas, start=2):
            col_letter = get_column_letter(j)
            c = ms.cell(
                row=total_row, column=j,
                value=f"=SUM({col_letter}{start_row + 1}:{col_letter}{start_row + len(largos)})"
            )
            c.font = total_font
            c.fill = total_fill
            c.border = border
            c.alignment = center

        col_start = get_column_letter(2)
        col_end = get_column_letter(total_col - 1)
        c = ms.cell(row=total_row, column=total_col, value=f"=SUM({col_start}{total_row}:{col_end}{total_row})")
        c.font = total_font
        c.fill = total_fill
        c.border = border
        c.alignment = center

        ms.column_dimensions["A"].width = 18
        for j in range(2, total_col + 1):
            ms.column_dimensions[get_column_letter(j)].width = 9

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ----------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# ----------------------------------------------------------------------------

st.set_page_config(page_title="Matriz de Producción desde PO", layout="wide")
st.title("📋 Generador de Matriz de Producción")
st.caption("Sube un PDF de orden de compra (formato Big R) y obtén la matriz Cintura x Largo por estilo/lavado.")

uploaded_file = st.file_uploader("Selecciona el archivo PDF de la orden de compra", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Leyendo y procesando el PDF..."):
        rows, meta, warnings = extract_lines_from_pdf(uploaded_file)

    if not rows:
        st.error(
            "No se pudo extraer ninguna línea de detalle del PDF. "
            "Es posible que el formato sea distinto al esperado."
        )
        st.stop()

    total_qty = sum(r["qty"] for r in rows)

    # --- Resumen general ---
    st.subheader("Resumen de la orden")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("PO #", meta.get("po_number", "N/D"))
    col2.metric("Líneas detectadas", len(rows))
    col3.metric("Total piezas", total_qty)
    col4.metric("Total piezas (según PDF)", meta.get("total_units_pdf", "N/D"))

    if meta.get("total_units_pdf") is not None and meta["total_units_pdf"] != total_qty:
        st.warning(
            f"⚠️ El total de piezas extraído ({total_qty}) no coincide con el total "
            f"declarado en el PDF ({meta['total_units_pdf']}). Revisa las advertencias abajo."
        )
    else:
        st.success("✅ El total de piezas coincide con el declarado en el PDF.")

    # --- Advertencias de líneas no interpretadas ---
    if warnings:
        with st.expander(f"⚠️ {len(warnings)} línea(s) no se pudieron interpretar automáticamente"):
            for w in warnings:
                st.text(w)

    # --- Inconsistencias descripción vs código MFG ---
    mismatches = detect_description_mfg_mismatches(rows)
    if mismatches:
        with st.expander(f"⚠️ {len(mismatches)} línea(s) con posible inconsistencia descripción/código MFG"):
            st.write(
                "Estas líneas tienen una descripción de estilo/lavado distinta a la dominante "
                "para su código MFG. Se agruparon por código MFG, pero conviene confirmar con el cliente."
            )
            for m in mismatches:
                st.text(
                    f"Línea {m['line_no']} | SKU {m['sku']} | Talla {m['size']} | "
                    f"MFG: {m['mfg']} | Descripción en línea: '{m['desc_en_linea']}' | "
                    f"Descripción esperada: '{m['desc_esperada_por_mfg']}'"
                )

    # --- Agrupación por estilo/lavado ---
    groups = group_by_mfg(rows)

    st.subheader("Vista previa de matrices por estilo / lavado")
    for mfg, g in groups.items():
        desc = max(g["descs"].items(), key=lambda kv: kv[1])[0]
        total = sum(r["qty"] for r in g["rows"])
        cinturas, largos, qty_map = build_size_matrix(g["rows"])

        st.markdown(f"**{desc}** — `{mfg}` — Total: {total} pzs")

        import pandas as pd
        table = pd.DataFrame(
            [[qty_map.get((c, l), 0) or "" for c in cinturas] for l in largos],
            index=[f"Largo {l}" for l in largos],
            columns=[f"Cintura {c}" for c in cinturas],
        )
        st.dataframe(table, use_container_width=True)

    # --- Generar y descargar Excel ---
    excel_bytes = build_excel(groups, meta, total_qty)

    st.subheader("Descargar resultado")
    st.download_button(
        label="⬇️ Descargar matriz en Excel",
        data=excel_bytes,
        file_name=f"Matriz_PO_{meta.get('po_number', 'orden')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Sube un PDF para comenzar.")
