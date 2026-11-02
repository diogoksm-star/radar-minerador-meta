# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class ExcelExporter:
    @staticmethod
    def export(offers_data: list, output_filepath: str):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Ofertas Validadas"

        header_fill_main = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
        header_fill_metodo_a = PatternFill(start_color="1E4D2B", end_color="1E4D2B", fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=9)
        link_font = Font(name="Calibri", size=9, color="0000FF", underline="single")
        
        thin_border = Border(
            left=Side(style='thin', color='E0E0E0'),
            right=Side(style='thin', color='E0E0E0'),
            top=Side(style='thin', color='E0E0E0'),
            bottom=Side(style='thin', color='E0E0E0')
        )
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

        headers = [
            "Rank", "Nicho / Setor", "Produto / Oferta", "Anunciante / Página",
            "Ads Ativos", "Longevidade", "Views Estimadas", "Gasto Estimado (R$)", 
            "Fat. Estimado (R$)", "Status Tração", "Preço Front", "Checkout", 
            "Link Landing Page", "Headline da LP",
            "[Método A] Gancho de Busca Google", "[Método A] Link do Post Facebook/Instagram",
            "[Método A] Views Reais no Player", "[Método A] Gasto Real Auditado (R$)"
        ]

        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.alignment = align_center
            if col_idx >= 15:
                cell.fill = header_fill_metodo_a
            else:
                cell.fill = header_fill_main

        for row_idx, o in enumerate(offers_data, start=2):
            m = o.get("metricas_thyerri", {})
            views_str = f"{m.get('views_estimadas', 0):,}".replace(",", ".")
            gasto_str = f"R$ {m.get('gasto_estimado_brl', 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            fat_str = f"R$ {m.get('faturamento_estimado_brl', 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            row = [
                f"#{row_idx - 1}",
                o.get("setor", "Geral"),
                o.get("nome", o.get("page_name", "")),
                o.get("page_name", ""),
                f"{o.get('ads_count', 1)} ads ativos",
                f"{o.get('max_days', 20)} dias no ar",
                views_str,
                gasto_str,
                fat_str,
                m.get("status_thyerri", "Validado"),
                o.get("preco", "R$ 29,90"),
                o.get("checkout", "Kiwify"),
                o.get("destination_lp", ""),
                o.get("headline", ""),
                o.get("metodo_a_gancho", ""),
                o.get("metodo_a_post_url", ""),
                o.get("metodo_a_views_reais", ""),
                o.get("metodo_a_gasto_real", "")
            ]
            ws.append(row)

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=c)
                cell.font = data_font
                cell.border = thin_border
                if c in [1, 5, 6, 10, 11, 12]:
                    cell.alignment = align_center
                elif c in [7, 8, 9, 17, 18]:
                    cell.alignment = align_right
                else:
                    cell.alignment = align_left

                if c in [13, 16] and cell.value and str(cell.value).startswith("http"):
                    cell.hyperlink = cell.value
                    cell.font = link_font

        # Auto width
        widths = {
            1: 8, 2: 24, 3: 30, 4: 22, 5: 14, 6: 15, 7: 18, 8: 20, 9: 20,
            10: 22, 11: 14, 12: 14, 13: 32, 14: 38, 15: 35, 16: 32, 17: 18, 18: 22
        }
        for col_i, w in widths.items():
            ws.column_dimensions[get_column_letter(col_i)].width = w

        wb.save(output_filepath)
        return output_filepath
