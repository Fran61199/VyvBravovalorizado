import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import io
import zipfile
import os
import tempfile

st.set_page_config(page_title="Separador de Preliquidación por Proyecto", layout="wide")
st.title("Separador de Preliquidación por Proyecto")
st.markdown("Sube un archivo de preliquidación y se generará un Excel por cada proyecto encontrado.")

uploaded_file = st.file_uploader("Sube el archivo de preliquidación (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active

    max_col = ws.max_column
    header_row = 4
    data_start_row = 5

    last_data_row = ws.max_row
    if ws.cell(row=ws.max_row, column=1).value == "TOTAL":
        last_data_row = ws.max_row - 1

    proyecto_col = 16

    proyectos = {}
    for row_idx in range(data_start_row, last_data_row + 1):
        proyecto = ws.cell(row=row_idx, column=proyecto_col).value
        if proyecto is None or str(proyecto).strip() == "":
            proyecto = "sin proyecto"
        if proyecto not in proyectos:
            proyectos[proyecto] = []
        proyectos[proyecto].append(row_idx)

    st.success(f"Se encontraron **{len(proyectos)}** proyectos en el archivo.")
    st.markdown("**Proyectos:**")
    for p, rows in sorted(proyectos.items(), key=lambda x: x[0]):
        st.markdown(f"- {p} ({len(rows)} atenciones)")

    if st.button("Generar Excels por Proyecto"):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for proyecto_name, row_indices in sorted(proyectos.items(), key=lambda x: x[0]):
                new_wb = openpyxl.Workbook()
                new_ws = new_wb.active
                new_ws.title = ws.title[:31] if len(ws.title) > 31 else ws.title

                for col in range(1, max_col + 1):
                    col_letter = get_column_letter(col)
                    if ws.column_dimensions[col_letter].width:
                        new_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width

                # Row 1 (empty, merged)
                new_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
                new_ws.row_dimensions[1].height = 15.75

                # Row 2 - Title with project name
                new_ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=17)
                title_cell = new_ws.cell(row=2, column=2, value=proyecto_name)
                title_cell.font = Font(name="Calibri", size=11, bold=True)
                title_cell.fill = PatternFill(start_color="FF00B050", end_color="FF00B050", fill_type="solid")
                title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                new_ws.row_dimensions[2].height = 30.0

                # Row 3 - CATALOGO header
                new_ws.merge_cells(start_row=3, start_column=19, end_row=3, end_column=max_col)
                catalogo_cell = new_ws.cell(row=3, column=19, value="CATALOGO")
                catalogo_cell.font = Font(name="Calibri", size=8, color="FFFFFFFF")
                catalogo_cell.fill = PatternFill(start_color="FF00B050", end_color="FF00B050", fill_type="solid")
                catalogo_cell.alignment = Alignment(horizontal="center")

                # Row 4 - Headers
                new_ws.row_dimensions[4].height = 28.5
                for col in range(1, max_col + 1):
                    src_cell = ws.cell(row=header_row, column=col)
                    dst_cell = new_ws.cell(row=4, column=col, value=src_cell.value)
                    dst_cell.font = Font(name="Calibri", size=8, bold=True, color="FFFFFFFF")
                    dst_cell.fill = PatternFill(start_color="FF80BFFF", end_color="FF80BFFF", fill_type="solid")
                    dst_cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                    dst_cell.border = Border(
                        left=Side(style="thin", color="FF000000"),
                        right=Side(style="thin", color="FF000000"),
                        top=Side(style="thin", color="FF000000"),
                        bottom=Side(style="thin", color="FF000000"),
                    )

                # Data rows
                new_row = 5
                for idx, src_row_idx in enumerate(row_indices, start=1):
                    new_ws.row_dimensions[new_row].height = 17.25
                    for col in range(1, max_col + 1):
                        src_cell = ws.cell(row=src_row_idx, column=col)
                        value = src_cell.value
                        if col == 1:
                            value = idx
                        dst_cell = new_ws.cell(row=new_row, column=col, value=value)
                        dst_cell.font = Font(name="Calibri", size=8)
                        dst_cell.alignment = Alignment(horizontal="center", wrap_text=True)
                        dst_cell.border = Border(
                            left=Side(style="thin", color="FF000000"),
                            right=Side(style="thin", color="FF000000"),
                            top=Side(style="thin", color="FF000000"),
                            bottom=Side(style="thin", color="FF000000"),
                        )
                        if src_cell.number_format:
                            dst_cell.number_format = src_cell.number_format
                    new_row += 1

                # Total row
                total_row = new_row
                new_ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=18)
                total_label_cell = new_ws.cell(row=total_row, column=1, value="TOTAL")
                total_label_cell.font = Font(name="Calibri", size=8, bold=True)
                total_label_cell.alignment = Alignment(horizontal="center")
                total_label_cell.border = Border(
                    left=Side(style="thin", color="FF000000"),
                    right=Side(style="thin", color="FF000000"),
                    top=Side(style="thin", color="FF000000"),
                    bottom=Side(style="thin", color="FF000000"),
                )

                for col in range(19, max_col + 1):
                    col_letter = get_column_letter(col)
                    start_cell = f"{col_letter}5"
                    end_cell = f"{col_letter}{total_row - 1}"
                    formula = f"=SUM({start_cell}:{end_cell})"
                    dst_cell = new_ws.cell(row=total_row, column=col, value=formula)
                    dst_cell.font = Font(name="Calibri", size=8, bold=True)
                    dst_cell.alignment = Alignment(horizontal="center")
                    dst_cell.border = Border(
                        left=Side(style="thin", color="FF000000"),
                        right=Side(style="thin", color="FF000000"),
                        top=Side(style="thin", color="FF000000"),
                        bottom=Side(style="thin", color="FF000000"),
                    )

                file_buffer = io.BytesIO()
                new_wb.save(file_buffer)
                file_buffer.seek(0)

                safe_name = proyecto_name.replace("/", "-").replace("\\", "-").replace(":", "-")
                filename = f"Preliquidacion_{safe_name}.xlsx"
                zf.writestr(filename, file_buffer.getvalue())

        zip_buffer.seek(0)
        st.download_button(
            label="Descargar todos los Excels (ZIP)",
            data=zip_buffer,
            file_name="Preliquidaciones_por_Proyecto.zip",
            mime="application/zip",
        )
        st.success("Archivos generados exitosamente.")
