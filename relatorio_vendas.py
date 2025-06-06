# Importa bibliotecas necessárias
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook  # Usado pelo pandas para escrever arquivos Excel
import seaborn as sns
import os

# ✅ Cria a pasta 'output' se ela não existir
os.makedirs('output', exist_ok=True)

# 📥 Lê os dados de vendas a partir do arquivo CSV
df = pd.read_csv('/content/vendas_simuladas.csv', sep=',')

# 🧮 Cria uma nova coluna chamada 'Receita' (Quantidade × Preço)
df['Receita'] = df['Quantidade'] * df['Preço']

# 📅 Cria uma coluna com ano e mês combinados (ex: 2024-03)
# Isso facilita a análise mensal
df['Data'] = pd.to_datetime(df['Data'])  # Corrige caso a coluna 'Data' esteja como string
df['AnoMes'] = df['Data'].dt.to_period('M')

# 📊 Calcula a receita total por mês
receita_mensal = df.groupby('AnoMes')['Receita'].sum().reset_index()

# 🥇 Calcula a quantidade total vendida por produto (ranking)
mais_vendidos = df.groupby('Produto')['Quantidade'].sum().sort_values(ascending=False)

# 💼 Calcula a receita total por vendedor
receita_vendedor = df.groupby('Vendedor')['Receita'].sum().sort_values(ascending=False)

# 🌍 Calcula a receita total por região
receita_regiao = df.groupby('Região')['Receita'].sum().sort_values(ascending=False)

# 📈 Cria gráfico da receita mensal
plt.figure(figsize=(8, 4))
sns.barplot(x=receita_mensal['AnoMes'].astype(str), y=receita_mensal['Receita'])  # Corrigido: "X" → "x"
plt.title('Receita por mês')
plt.ylabel('Receita (R$)')
plt.xlabel('Ano-Mês')
plt.tight_layout()

# 💾 Salva o gráfico como imagem
plt.savefig('output/receita_mensal.png')
plt.close()

# 📤 Exporta os dados e análises para um arquivo Excel com múltiplas abas
with pd.ExcelWriter('output/relatorio_final.xlsx') as writer:
    df.to_excel(writer, sheet_name='Base de Dados', index=False)
    receita_mensal.to_excel(writer, sheet_name='Receita Mensal', index=False)
    mais_vendidos.to_frame('Quantidade').to_excel(writer, sheet_name='Mais Vendidos')
    receita_vendedor.to_frame('Receita').to_excel(writer, sheet_name='Receita por Vendedor')
    receita_regiao.to_frame('Receita').to_excel(writer, sheet_name='Receita por Região')

# ✅ Mensagem de confirmação
print("Relatório gerado em: output/relatorio_final.xlsx")
