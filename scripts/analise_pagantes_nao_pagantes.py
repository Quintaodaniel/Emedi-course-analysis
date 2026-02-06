import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import textwrap

# --- 1. CONFIGURAÇÕES E DADOS ---
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
df = pd.read_excel(caminho_xlsx, sheet_name=0)

# Colunas necessárias
colunas = [
    'nome_curso', 
    'pagantes_inscritos', 'pagantes_confirmados',
    'nao_pagantes_inscritos', 'nao_pagantes_confirmados'
]
df = df[colunas].dropna(subset=['nome_curso'])

# Preencher Nulos
for col in colunas[1:]:
    df[col] = df[col].fillna(0).astype(int)

# --- 2. TRATAMENTO DOS NOMES ---
labels_com_quebra = [textwrap.fill(nome, width=20) for nome in df['nome_curso']]

# --- 3. CONSTRUÇÃO DO GRÁFICO ---
x = np.arange(len(labels_com_quebra))
width = 0.20  # Diminuímos a largura para caber 4 barras por curso

fig, ax = plt.subplots(figsize=(18, 9))

# Plotagem das 4 barras por curso
rects1 = ax.bar(x - width*1.5, df['pagantes_inscritos'], width, label='Inscritos (Pagantes)', color='#e67e22', zorder=3)
rects2 = ax.bar(x - width*0.5, df['pagantes_confirmados'], width, label='Confirmados (Pagantes)', color='#f1c40f', zorder=3)
rects3 = ax.bar(x + width*0.5, df['nao_pagantes_inscritos'], width, label='Inscritos (Não Pagantes)', color='#9b59b6', zorder=3)
rects4 = ax.bar(x + width*1.5, df['nao_pagantes_confirmados'], width, label='Confirmados (Não Pagantes)', color='#95a5a6', zorder=3)

# --- 4. ESTÉTICA E RÓTULOS ---
ax.set_ylabel('Quantidade de Pessoas', fontweight='bold', fontsize=12)
ax.set_title('Visão Geral: Inscritos vs Confirmados (Pagantes e Não Pagantes)', fontsize=20, fontweight='bold', pad=30)

ax.set_xticks(x)
ax.set_xticklabels(labels_com_quebra, fontsize=10, ha='center') 

# Legenda organizada em 2 colunas para poupar espaço
ax.legend(loc='upper right', frameon=True, shadow=True, ncol=2)
ax.grid(axis='y', linestyle='--', alpha=0.6, zorder=0)

# Função para colocar os números sobre as barras
def autolabel(rects):
    for rect in rects:
        height = rect.get_height()
        if height > 0: # Só mostra se for maior que zero para não poluir
            ax.annotate(f'{int(height)}',
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 5), textcoords="offset points",
                        ha='center', va='bottom', fontweight='bold', fontsize=9)

autolabel(rects1)
autolabel(rects2)
autolabel(rects3)
autolabel(rects4)

plt.tight_layout()

# --- 5. SALVAR ---
pasta_output = 'outputs'
if not os.path.exists(pasta_output): os.makedirs(pasta_output)

caminho_png = os.path.join(pasta_output, 'grafico_pagantes_nao_pagantes.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print(f"Sucesso! Gráfico consolidado gerado: {caminho_png}")
plt.show()