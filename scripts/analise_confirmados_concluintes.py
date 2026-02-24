import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import textwrap

# --- 1. CONFIGURAÇÕES E DADOS ---
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
df = pd.read_excel(caminho_xlsx, sheet_name=0)

# --- 2. FILTRO DE CURSOS ---
# Coloque aqui os IDs dos cursos que você quer que apareçam no gráfico
# Exemplo: ids_selecionados = ['C1', 'C2', 'C5', 'C6']
# Se quiser ver TODOS, basta comentar a linha do filtro abaixo.
ids_selecionados = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6'] 

# Selecionando colunas necessárias (incluindo o ID para o filtro)
df = df[['id', 'nome_curso', 'qtd_confirmados', 'qtd_concluintes']]

# Aplicando o filtro (se a lista não estiver vazia)
if ids_selecionados:
    df = df[df['id'].isin(ids_selecionados)]

# Limpeza básica
df = df.dropna(subset=['nome_curso'])
df['qtd_confirmados'] = df['qtd_confirmados'].fillna(0).astype(int)
df['qtd_concluintes'] = df['qtd_concluintes'].fillna(0).astype(int)

# --- 3. TRATAMENTO DOS NOMES ---
labels_com_quebra = [textwrap.fill(nome, width=20) for nome in df['nome_curso']]

# --- 4. CONSTRUÇÃO DO GRÁFICO ---
x = np.arange(len(labels_com_quebra))
width = 0.35

fig, ax = plt.subplots(figsize=(15, 8))

# Cores: Verde (Confirmados) e Azul Marinho (Concluintes)
rects1 = ax.bar(x - width/2, df['qtd_confirmados'], width, label='Confirmados', color='#2ecc71', zorder=3)
rects2 = ax.bar(x + width/2, df['qtd_concluintes'], width, label='Concluintes', color='#34495e', zorder=3)

# --- 5. ESTÉTICA E RÓTULOS ---
ax.set_ylabel('Quantidade de Alunos', fontweight='bold', fontsize=12)
ax.set_title('Relação: Confirmados vs Concluintes', fontsize=18, fontweight='bold', pad=30)

ax.set_xticks(x)
ax.set_xticklabels(labels_com_quebra, fontsize=10, ha='center') 

ax.legend(loc='upper right', frameon=True, shadow=True)
ax.grid(axis='y', linestyle='--', alpha=0.6, zorder=0)

def autolabel(rects):
    for rect in rects:
        height = rect.get_height()
        ax.annotate(f'{int(height)}',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 5), textcoords="offset points",
                    ha='center', va='bottom', fontweight='bold', fontsize=10)

autolabel(rects1)
autolabel(rects2)

plt.tight_layout()

# --- 6. SALVAR ---
pasta_output = 'outputs'
if not os.path.exists(pasta_output): os.makedirs(pasta_output)

caminho_png = os.path.join(pasta_output, 'grafico_confirmados_concluintes.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print(f"Sucesso! Gráfico gerado para {len(df)} cursos em: {caminho_png}")
plt.show()