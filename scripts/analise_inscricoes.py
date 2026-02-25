import pandas as pd
import matplotlib.pyplot as plt
import os
from adjustText import adjust_text 

# Configurações de estilo para um visual moderno
plt.style.use('seaborn-v0_8-whitegrid')

# Caminhos
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
pasta_output = 'outputs'
if not os.path.exists(pasta_output):
    os.makedirs(pasta_output)

# --- 1. CARREGAR E PREPARAR DADOS ---
df_cursos_ref = pd.read_excel(caminho_xlsx, sheet_name=0)
df_eventos_ref = pd.read_excel(caminho_xlsx, sheet_name=1)
df_dados = pd.read_excel(caminho_xlsx, sheet_name=3)

df_nomes_cursos = df_cursos_ref[['id', 'nome_curso']].rename(columns={'nome_curso': 'nome_exibicao'})
df_nomes_eventos = df_eventos_ref[['id', 'nome_evento']].rename(columns={'nome_evento': 'nome_exibicao'})
df_referencia = pd.concat([df_nomes_cursos, df_nomes_eventos], ignore_index=True)

df_dados['data'] = pd.to_datetime(df_dados['data'])
df_final = pd.merge(df_dados, df_referencia, on='id', how='left')

# Filtro dos itens desejados
ids_para_grafico = ['C8', 'E1', 'C9', 'C10', 'C11', 'E2'] 
df_filtrado = df_final[df_final['id'].isin(ids_para_grafico)].sort_values('data')
df_filtrado['data_str'] = df_filtrado['data'].dt.strftime('%d/%m/%Y')

# --- 2. CONSTRUÇÃO DO GRÁFICO ---
fig, ax = plt.subplots(figsize=(16, 8))

# Lista para os rótulos
textos = []

# Cores mais elegantes (opcional: o matplotlib escolherá automaticamente se preferir)
cores = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']

for i, id_item in enumerate(ids_para_grafico):
    dados_item = df_filtrado[df_filtrado['id'] == id_item]
    
    if not dados_item.empty:
        nome_texto = dados_item['nome_exibicao'].iloc[0]
        label_legenda = f"{id_item} - {nome_texto}"
        
        # Desenhar a linha com sombra suave e marcadores destacados
        linha = ax.plot(dados_item['data_str'], dados_item['qtd_inscritos'], 
                        marker='o', markersize=8, markeredgecolor='white', 
                        markeredgewidth=1.5, linewidth=3, label=label_legenda,
                        alpha=0.9)
        
        cor_da_linha = linha[0].get_color()
        
        # Adicionar os números
        for x, y in zip(dados_item['data_str'], dados_item['qtd_inscritos']):
            txt = ax.text(x, y, f' {int(y)} ', 
                          fontsize=10, 
                          fontweight='bold', 
                          color=cor_da_linha,
                          # Fundo branco para não misturar com as linhas
                          bbox=dict(facecolor='white', alpha=0.8, edgecolor='none', pad=0.5))
            textos.append(txt)

# --- 3. AJUSTE AUTOMÁTICO DOS RÓTULOS ---
adjust_text(textos, 
            ax=ax,
            # Força o movimento apenas no eixo Y para manter o alinhamento com a data
            only_move={'points':'y', 'text':'y'}, 
            # Adiciona linhas guia discretas se o texto se afastar muito
            arrowprops=dict(arrowstyle='-', color='gray', lw=0.5, alpha=0.5),
            expand_text=(1.2, 1.4),
            expand_points=(1.2, 1.4))

# --- 4. PERFUMARIA E ACABAMENTO ---
ax.set_title('Evolução de Inscritos por Curso/Evento', fontsize=18, fontweight='bold', pad=30, loc='left', color='#333333')
ax.set_ylabel('Total de Inscritos', fontsize=12, color='#666666')
ax.set_xlabel('Data da Verificação', fontsize=12, labelpad=10, color='#666666')

# Rotacionar datas
plt.xticks(rotation=45, ha='right')

# Limpar bordas do gráfico (deixar apenas o necessário)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Melhorar o grid
ax.grid(axis='y', linestyle='--', alpha=0.4)
ax.grid(axis='x', visible=False)

# Legenda em 2 colunas para ocupar menos espaço vertical
ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.18), 
          frameon=True, shadow=False, ncol=2, fontsize=10)

# Ajuste de limite superior para dar "ar" ao gráfico
if not df_filtrado.empty:
    ax.set_ylim(bottom=-5, top=df_filtrado['qtd_inscritos'].max() * 1.1)

plt.subplots_adjust(bottom=0.25)

# Salvar
caminho_png = os.path.join(pasta_output, 'relatorio_inscritos.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print("Gráfico gerado com sucesso!")
plt.show()