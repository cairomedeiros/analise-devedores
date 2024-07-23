import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox

def carregar_dados(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    return df

def analisar_fiado(df):
    fiado_df = df[df['Fiado'] == 'Sim']
    total_fiado_cliente = fiado_df.groupby('Cliente')['Receita Total'].sum()
    total_fiado_produto = fiado_df.groupby('Produto')['Receita Total'].sum()
    return total_fiado_cliente, total_fiado_produto

def visualizar_dados(total_fiado_cliente, total_fiado_produto):
    # Gráfico de vendas fiado por cliente
    total_fiado_cliente.plot(kind='bar', color='skyblue')
    plt.title('Total de Vendas Fiado por Cliente')
    plt.xlabel('Cliente')
    plt.ylabel('Receita Fiado (R$)')
    plt.show()
    plt.close()  # Fecha o gráfico após exibi-lo

    # Gráfico de vendas fiado por produto
    total_fiado_produto.plot(kind='bar', color='lightgreen')
    plt.title('Total de Vendas Fiado por Produto')
    plt.xlabel('Produto')
    plt.ylabel('Receita Fiado (R$)')
    plt.show()
    plt.close()  # Fecha o gráfico após exibi-lo

    # Criação da string para exibir os valores
    total_fiado_str = "Total de Vendas Fiado por Cliente:\n"
    for cliente, valor in total_fiado_cliente.items():
        total_fiado_str += f"{cliente}: R$ {valor:.2f}\n"
    
    return total_fiado_str

def carregar_e_analisar():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if caminho_arquivo:
        try:
            df = carregar_dados(caminho_arquivo)
            total_fiado_cliente, total_fiado_produto = analisar_fiado(df)
            total_fiado_str = visualizar_dados(total_fiado_cliente, total_fiado_produto)
            lbl_resultado.config(text=total_fiado_str)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {e}")

root = tk.Tk()
root.title("Análise de Vendas Fiado")

btn_carregar = tk.Button(root, text="Carregar Arquivo Excel e Analisar", command=carregar_e_analisar)
btn_carregar.pack(pady=20)

lbl_resultado = tk.Label(root, text="")
lbl_resultado.pack(pady=20)

root.mainloop()


