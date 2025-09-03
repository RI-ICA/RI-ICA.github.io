import pandas as pd
import urllib.parse
import svgwrite

# --- Parâmetros para o acesso à planilha ---
sheet_id = "1UiD53PypgwavUi8r6YzpuAFZ7XsItJAk_siFjNU4Rq4"
sheet_name = "2024-2025"
encoded_sheet_name = urllib.parse.quote(sheet_name)
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_sheet_name}"

# --- Ler os dados ---
df = pd.read_csv(url) 
df = df.fillna('')  # substitui NaN por string vazia
df = df[['SEM', 'Type', 'CODE', 'Nom Francais', 'ECTS', 'hours in class']]

# --- Garantir tipos de dados corretos ---
df['ECTS'] = df['ECTS'].astype(str)
df['hours in class'] = df['hours in class'].astype(str)
df['Type'] = df['Type'].astype(str)

# --- semesters
df_s1 = df[df['SEM'] == 'S1'].reset_index(drop=True)
df_s2 = df[df['SEM'] == 'S2'].reset_index(drop=True)
df_s3 = df[df['SEM'] == 'S3'].reset_index(drop=True)
df_s4 = df[df['SEM'] == 'S4'].reset_index(drop=True)
df_s5 = df[df['SEM'] == 'S5'].reset_index(drop=True)
df_s6 = df[df['SEM'] == 'S6'].reset_index(drop=True)
df_s7 = df[df['SEM'] == 'S7'].reset_index(drop=True)
df_s8 = df[df['SEM'] == 'S8'].reset_index(drop=True)
df_s9 = df[df['SEM'] == 'S9'].reset_index(drop=True)


#print(df.columns)
#print(df[['SEM', 'Type', 'CODE', 'Nom Francais', 'ECTS', 'hours in class']])
#print(df_s3.to_string())

# --- Parâmetros gráficos ---
rect_width = 120
rect_height = 50
y_spacing = 20
x_offset = 40
y_offset = 40
font_size = 6

# Parâmetros de layout
col_gap = 25  # espaço entre colunas
x_s1 = x_offset
x_s2 = x_s1 + rect_width + col_gap
x_s3 = x_s2 + rect_width + col_gap
x_s4 = x_s3 + rect_width + col_gap
x_s5 = x_s4 + rect_width + col_gap
x_s6 = x_s5 + rect_width + col_gap
x_s7 = x_s6 + rect_width + col_gap
x_s8 = x_s7 + rect_width + col_gap
x_s9 = x_s8 + rect_width + col_gap

# --- Cores conforme tipo ---
colors = {
    'UNITE': "#D29B35",        # marrom escuro
    'COMPOSANTE': "#fff4b2",   # amarelo claro
    'default': "#ccffd7"       # verde claro
}

def draw_col(dwg, df_act, x_act, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo=None):
    # Adicionar título da coluna, se fornecido
    if titulo:
        dwg.add(dwg.text(titulo, insert=(x_act, y_offset - 30),
                         font_size=14, font_weight='bold', fill='black'))

    for j, row in df_act.iterrows():
        y = y_offset + j * (rect_height + y_spacing)
        type_str = row['Type'].strip().upper()
        fill_color = colors.get(type_str, colors['default'])

        # Retângulo
        dwg.add(dwg.rect(insert=(x_act, y), size=(rect_width, rect_height),
                         stroke='black', fill=fill_color, stroke_width=0.5))

        # Texto interno
        lines = [
            str(row['CODE']),
            str(row['Nom Francais']),
            f"ECTS: {row['ECTS']}, Hours: {row['hours in class']}"
        ]
        for k, line in enumerate(lines):
            dwg.add(dwg.text(line, insert=(x_act + 5, y + 15 + k * 12),
                             font_size=font_size, fill='black'))


# --- Criar desenho SVG ---
dwg = svgwrite.Drawing('matrice_semesters5.svg', size=('400mm', '700mm'))

# --- Título do semestre ---
#dwg.add(dwg.text("SEM 3", insert=(x_s3, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 4", insert=(x_s4, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 5", insert=(x_s5, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 6", insert=(x_s6, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 7", insert=(x_s7, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 8", insert=(x_s8, y_offset - 10), font_size=12, font_weight='bold', fill='black'))
#dwg.add(dwg.text("SEM 9", insert=(x_s9, y_offset - 10), font_size=12, font_weight='bold', fill='black'))


draw_col(dwg, df_s1, x_s1, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 1")
draw_col(dwg, df_s2, x_s2, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 2")
draw_col(dwg, df_s3, x_s3, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 3")
draw_col(dwg, df_s4, x_s4, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 4")
draw_col(dwg, df_s5, x_s5, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 5")
draw_col(dwg, df_s6, x_s6, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 6")
draw_col(dwg, df_s7, x_s7, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 7")
draw_col(dwg, df_s8, x_s8, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 8")
draw_col(dwg, df_s9, x_s9, y_offset, rect_width, rect_height, y_spacing, font_size, colors, titulo="SEM 9")

# Coordenadas dos centros verticais dos primeiros retângulos
y_s3 = y_offset + rect_height / 2
x_end_s3 = x_s3 + rect_width

y_s4 = y_offset + rect_height / 2
x_start_s4 = x_s4

marker = dwg.marker(insert=(10, 5), size=(10, 10), orient="auto")
marker.add(dwg.path(d="M0,0 L10,5 L0,10 L2,5 Z", fill="black"))
dwg.defs.add(marker)

# desenhar seta entre os primeiros retângulos
y1 = y_offset + rect_height / 2
dwg.add(dwg.line(start=(x_s3 + rect_width, y1), end=(x_s4, y1),
                                   stroke='black', stroke_width=1, marker_end=marker.get_funciri()))



# --- Salvar arquivo ---
dwg.save()
print("✅ SVG gerado com sucesso!")