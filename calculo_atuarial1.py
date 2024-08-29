import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Definir os dados
data = {
    'x': [20, 21, 22, 23, 24, 25],
    'lx': [181.524, 181.179, 180.835, 179.458, 179.120, 178.855],
    'i': 0.05  # Taxa de juros de 6%
}

# Converter os dados para um DataFrame
df = pd.DataFrame(data)

# Calcular vx
df['vx'] = (1 + df['i']) ** -(df['x'] - df['x'][0])

# Calcular Dx
df['DX'] = df['lx'] * df['vx']

# Criar o arquivo Excel e escrever os dados #adicionar caminho aqui
with pd.ExcelWriter('D:/Python_Lib/vin/valores_atuariais.xlsx',
                    engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Dados', index=False)
    # Gráfico de vx
    plt.figure(figsize=(10, 6))
    plt.plot(df['x'], df['vx'], marker='o', color='blue')
    plt.title('Gráfico de vx em função de x')
    plt.xlabel('Idade (x)')
    plt.ylabel('vx')
    plt.grid(True)
    plt.savefig(
        'D:/Python_Lib/vin/vx_plot.png')  # Salvar o gráfico como uma imagem
    # Adicionar o gráfico ao Excel
    worksheet = writer.sheets['Dados']
    img = Image('D:/Python_Lib/vin/vx_plot.png')
    worksheet.add_image(img, 'H2')
    # Gráfico de DX
    plt.figure(figsize=(10, 6))
    plt.plot(df['x'], df['DX'], marker='o', color='green')
    plt.title('Gráfico de DX em função de x')
    plt.xlabel('Idade (x)')
    plt.ylabel('DX')
    plt.grid(True)
    plt.savefig(
        'D:/Python_Lib/vin/vx_plot.png')  # Salvar o gráfico como uma imagem
    # Adicionar o gráfico ao Excel
    img = Image('D:/Python_Lib/vin/vx_plot.png')
    worksheet.add_image(img, 'H20')

'D:/Python_Lib/vin/valores_atuariais.xlsx'

#definir dados
# data = {
#     'x': []
#     'Lx': [],
#     'i': 0,06 #taxa de juros de 6%
# }
#
# df = pd.DataFrame