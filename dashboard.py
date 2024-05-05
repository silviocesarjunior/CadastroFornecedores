import matplotlib.pyplot as plt
import openpyxl

# Carregar dados do Excel
wb = openpyxl.load_workbook('Entradas.xlsx')
sheet = wb.active

# Inicializar dicionário para contar fornecedores e armazenar as horas de cada fornecedor
fornecedores = {}
horas_por_fornecedor = {}

# Percorrer as linhas e somar as horas para cada fornecedor
for row in sheet.iter_rows(min_row=2, values_only=True):
    fornecedor = row[1]
    horas = row[8]  # Considerando a coluna "Diferença" como a nona coluna (índice 8)
    if fornecedor in fornecedores:
        fornecedores[fornecedor] += 1
        if fornecedor in horas_por_fornecedor:
            horas_por_fornecedor[fornecedor].append(horas)  # Armazenar as horas como strings
        else:
            horas_por_fornecedor[fornecedor] = [horas]  # Armazenar as horas como strings
    else:
        fornecedores[fornecedor] = 1
        horas_por_fornecedor[fornecedor] = [horas]  # Armazenar as horas como strings

# Calcular a soma das horas e a média para cada fornecedor
soma_horas_por_fornecedor = {
    fornecedor: sum(int(hora[:2]) * 60 + int(hora[3:]) for hora in horas) / 60  # Calcular a soma das horas convertendo para minutos e, em seguida, dividindo por 60 para obter o total em horas
    for fornecedor, horas in horas_por_fornecedor.items()
}
media_horas_por_fornecedor = {
    fornecedor: soma_horas / fornecedores[fornecedor]  # Calcular a média dividindo a soma das horas pela quantidade de ocorrências do fornecedor
    for fornecedor, soma_horas in soma_horas_por_fornecedor.items()
}

# Preparar dados para o gráfico de barras
fornecedores_lista = list(fornecedores.keys())
quantidades = list(fornecedores.values())

# Criar o gráfico de barras
plt.figure(figsize=(10, 6))
bars = plt.bar(fornecedores_lista, quantidades, color='teal')

# Adicionar rótulos com a quantidade e a média de horas em cada barra
for bar, media_horas in zip(bars, media_horas_por_fornecedor.values()):
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2.0, height, f'{int(height)}', ha='center', va='bottom')
    plt.text(bar.get_x() + bar.get_width()/2.0, height/2, f'Média: {media_horas:.2f} horas', ha='center', va='bottom')

plt.xlabel('\n\nFornecedores')
plt.ylabel('Quantidade de pedidos')
plt.title('Quantidade de Registros e Média de Horas por Fornecedor')
plt.xticks(rotation=45, ha='right')
plt.yticks(range(1, max(quantidades) + 1))  # Ajuste para mostrar valores inteiros no eixo y de 1 em 1
plt.tight_layout()

# Adicionar botão para salvar em PDF
plt.savefig('grafico_fornecedores.pdf')

# Mostrar o gráfico
plt.show()
