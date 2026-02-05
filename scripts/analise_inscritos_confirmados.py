import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import textwrap  # Biblioteca para quebrar o texto

# --- 1. CONFIGURAÇÕES E DADOS ---
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
df = pd.read_excel(caminho_xlsx, sheet_name=0)
df = df[['nome_curso', 'total_inscritos', 'qtd_confirmados']].dropna(subset=['nome_curso'])

df['total_inscritos'] = df['total_inscritos'].fillna(0).astype(int)
df['qtd_confirmados'] = df['qtd_confirmados'].fillna(0).astype(int)

# --- 2. TRATAMENTO DOS NOMES (QUEBRA DE LINHA) ---
# Aqui definimos que cada linha terá no máximo 20 caracteres
labels_com_quebra = [textwrap.fill(nome, width=20) for nome in df['nome_curso']]

# --- 3. CONSTRUÇÃO DO GRÁFICO ---
x = np.arange(len(labels_com_quebra))
width = 0.35

fig, ax = plt.subplots(figsize=(15, 8)) # Ajustamos a largura para caber os textos

rects1 = ax.bar(x - width/2, df['total_inscritos'], width, label='Total Inscritos', color='#3498db', zorder=3)
rects2 = ax.bar(x + width/2, df['qtd_confirmados'], width, label='Confirmados', color='#2ecc71', zorder=3)

# --- 4. ESTÉTICA E RÓTULOS ---
ax.set_ylabel('Quantidade de Pessoas', fontweight='bold', fontsize=12)
ax.set_title('Inscrições por Curso: Total vs Confirmados', fontsize=18, fontweight='bold', pad=30)

# Aplicando os nomes com quebra de linha no eixo X
ax.set_xticks(x)
ax.set_xticklabels(labels_com_quebra, fontsize=10, ha='center') 

ax.legend(loc='upper right', frameon=True, shadow=True)
ax.grid(axis='y', linestyle='--', alpha=0.6, zorder=0)

# Função para colocar os números sobre as barras
def autolabel(rects):
    for rect in rects:
        height = rect.get_height()
        ax.annotate(f'{int(height)}',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 5), textcoords="offset points",
                    ha='center', va='bottom', fontweight='bold', fontsize=10)

autolabel(rects1)
autolabel(rects2)

# Ajuste de layout para não cortar os nomes embaixo
plt.tight_layout()

# --- 5. SALVAR ---
pasta_output = 'outputs'
if not os.path.exists(pasta_output): os.makedirs(pasta_output)
caminho_png = os.path.join(pasta_output, 'grafico_inscricoes_cancelados.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print(f"Sucesso! Gráfico gerado com nomes formatados: {caminho_png}")
plt.show()