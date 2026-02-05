import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os

# Caminhos
caminho_xlsx = 'data/planilha_geral_cursos.xlsx'
pasta_output = 'outputs'

if not os.path.exists(pasta_output):
    os.makedirs(pasta_output)

# --- 1. CARREGAR OS DADOS ---
df_cursos_ref = pd.read_excel(caminho_xlsx, sheet_name=0)   # Cadastro de Cursos
df_eventos_ref = pd.read_excel(caminho_xlsx, sheet_name=1)  # Eventos
df_dados = pd.read_excel(caminho_xlsx, sheet_name=3)        # Inscricoes por dia

# --- 2. CONSOLIDAR NOMES (Para a Legenda) ---
# Padronizamos a coluna de nome para 'nome_exibicao' em ambas as tabelas
df_nomes_cursos = df_cursos_ref[['id', 'nome_curso']].rename(columns={'nome_curso': 'nome_exibicao'})
df_nomes_eventos = df_eventos_ref[['id', 'nome_evento']].rename(columns={'nome_evento': 'nome_exibicao'})

# Juntamos as duas tabelas de referência
df_referencia = pd.concat([df_nomes_cursos, df_nomes_eventos], ignore_index=True)

# --- 3. MERGE E FILTRO ---
df_dados['data'] = pd.to_datetime(df_dados['data'])
df_final = pd.merge(df_dados, df_referencia, on='id', how='left')

# ADICIONADO O 'E1' NA LISTA ABAIXO:
ids_para_grafico = ['C7', 'C8', 'E1'] 

df_filtrado = df_final[df_final['id'].isin(ids_para_grafico)].sort_values('data')

# --- 4. CONSTRUÇÃO DO GRÁFICO ---
plt.figure(figsize=(14, 8))
ax = plt.gca()

for id_item in ids_para_grafico:
    dados_item = df_filtrado[df_filtrado['id'] == id_item]
    
    if not dados_item.empty:
        # Pega o nome vindo do merge (ou o ID caso o nome esteja vazio)
        nome_texto = dados_item['nome_exibicao'].iloc[0]
        label_legenda = f"{id_item} - {nome_texto}"
        
        # Plotar
        linha = plt.plot(dados_item['data'], dados_item['qtd_inscritos'], 
                         marker='o', linewidth=2, label=label_legenda)
        
        # Rótulos de dados (números sobre os pontos)
        for x, y in zip(dados_item['data'], dados_item['qtd_inscritos']):
            plt.text(x, y + 1, f'{int(y)}', ha='center', va='bottom', 
                     fontsize=9, fontweight='bold', color=linha[0].get_color())

# FORMATAR EIXO X
ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))

# Customizações Visuais
plt.title('Evolução Diária de Inscritos (Cursos e Eventos)', fontsize=16, fontweight='bold', pad=20)
plt.xlabel('Data da Verificação', fontsize=12, labelpad=15)
plt.ylabel('Total de Inscritos', fontsize=12)
plt.xticks(rotation=45)
plt.grid(True, linestyle='--', alpha=0.5)

# Legenda: ajustada para acomodar mais itens se necessário
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), 
           shadow=True, ncol=1, fontsize=9)

plt.subplots_adjust(bottom=0.3) 

# Salvar
caminho_png = os.path.join(pasta_output, 'relatorio_inscritos.png')
plt.savefig(caminho_png, dpi=300, bbox_inches='tight')

print(f"Sucesso! Gráfico gerado para: {', '.join(ids_para_grafico)}")
plt.show()