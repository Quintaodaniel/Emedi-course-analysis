import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import textwrap

# --- 1. CONFIGURAÇÕES E DADOS ---
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
df = pd.read_excel(caminho_xlsx, sheet_name=0)

# --- 2. FILTRO DE CURSOS ---
ids_selecionados = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6'] 

# Adicionamos 'qtd_nunca_presentes' na seleção de colunas
colunas_necessarias = ['id', 'nome_curso', 'qtd_confirmados', 'qtd_concluintes', 'qtd_nunca_presentes']
df = df[colunas_necessarias]

# Aplicando o filtro
if ids_selecionados:
    df = df[df['id'].isin(ids_selecionados)]

# Limpeza básica e conversão para numérico
df = df.dropna(subset=['nome_curso'])
for col in ['qtd_confirmados', 'qtd_concluintes', 'qtd_nunca_presentes']:
    df[col] = df[col].fillna(0).astype(int)

# --- NOVO CÁLCULO ---
# Pessoas que assistiram pelo menos 1 aula = Confirmados - Nunca Presentes
df['qtd_assistiram'] = df['qtd_confirmados'] - df['qtd_nunca_presentes']

# --- 3. TRATAMENTO DOS NOMES ---
labels_com_quebra = [textwrap.fill(nome, width=20) for nome in df['nome_curso']]

# --- 4. CONSTRUÇÃO DO GRÁFICO ---
x = np.arange(len(labels_com_quebra))
width = 0.25  # Diminuímos a largura para caberem 3 barras

fig, ax = plt.subplots(figsize=(16, 8))

# Cores: Verde (Confirmados), Azul Claro (Assistiram) e Azul Escuro (Concluintes)
rects1 = ax.bar(x - width, df['qtd_confirmados'], width, label='Confirmados', color='#2ecc71', zorder=3)
rects2 = ax.bar(x, df['qtd_assistiram'], width, label='Assistiram (mín. 1 aula)', color='#3498db', zorder=3)
rects3 = ax.bar(x + width, df['qtd_concluintes'], width, label='Concluintes', color='#34495e', zorder=3)

# --- 5. ESTÉTICA E RÓTULOS ---
ax.set_ylabel('Quantidade de Alunos', fontweight='bold', fontsize=12)
ax.set_title('Relação: Confirmados vs Assistiram vs Concluintes', fontsize=18, fontweight='bold', pad=30)

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
                    ha='center', va='bottom', fontweight='bold', fontsize=9)

autolabel(rects1)
autolabel(rects2)
autolabel(rects3)

plt.tight_layout()

# --- 6. SALVAR ---
pasta_output = 'outputs'
if not os.path.exists(pasta_output): os.makedirs(pasta_output)

caminho_png = os.path.join(pasta_output, 'grafico_confirmados_concluintes.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print(f"Sucesso! Gráfico gerado com a nova coluna em: {caminho_png}")
plt.show()