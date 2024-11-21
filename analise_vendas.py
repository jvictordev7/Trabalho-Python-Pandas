import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configuração do estilo dos gráficos
sns.set_theme(style="whitegrid")

# Leitura do arquivo
arquivo_vendas = "PlanilhaVendas.xlsx"
df = pd.read_excel(arquivo_vendas)

# Calcula a receita total por venda
df['Receita Total'] = df['Quantidade'] * df['Preço unitário']

# Produto mais vendido
produto_mais_vendido = df.groupby('ID do produto')['Quantidade'].sum().idxmax()
quantidade_vendida = df.groupby('ID do produto')['Quantidade'].sum().max()

# Receita total por região
receita_por_regiao = df.groupby('Região')['Receita Total'].sum()

# Dia com mais vendas
dia_com_mais_vendas = df['Data da venda'].value_counts().idxmax()
total_vendas_dia = df['Data da venda'].value_counts().max()

# Geração do relatório
relatorio = df.groupby('ID do produto').agg({
    'Quantidade': 'sum',
    'Receita Total': 'sum'
}).reset_index()

relatorio.rename(columns={
    'Quantidade': 'Quantidade Total Vendida',
    'Receita Total': 'Receita Total'
}, inplace=True)

relatorio = relatorio.sort_values(by='Receita Total', ascending=False)

# Salvar o relatório
arquivo_relatorio = "relatorio_vendas.xlsx"
relatorio.to_excel(arquivo_relatorio, index=False)

# Visualização 1: Receita por região
plt.figure(figsize=(8, 6))
receita_por_regiao.plot(kind='bar', color=sns.color_palette("coolwarm", len(receita_por_regiao)))
plt.title("Receita Total por Região", fontsize=16)
plt.ylabel("Receita Total (R$)")
plt.xlabel("Região")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("receita_por_regiao.png")  # Salva o gráfico como imagem
plt.show()

# Visualização 2: Produtos mais vendidos
plt.figure(figsize=(8, 6))
sns.barplot(
    data=relatorio.head(5),
    x="Quantidade Total Vendida",
    y="ID do produto",
    palette="viridis"
)
plt.title("Top 5 Produtos Mais Vendidos", fontsize=16)
plt.xlabel("Quantidade Total Vendida")
plt.ylabel("ID do Produto")
plt.tight_layout()
plt.savefig("top_produtos.png")
plt.show()

# Resultados no terminal
print(f"Produto mais vendido: {produto_mais_vendido} com {quantidade_vendida} unidades.")
print("Receita por região:")
print(receita_por_regiao)
print(f"Dia com mais vendas: {dia_com_mais_vendas} com {total_vendas_dia} vendas.")
print(f"Relatório salvo em: {arquivo_relatorio}")
