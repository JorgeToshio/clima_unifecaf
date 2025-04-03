import tkinter as tk
from openpyxl import load_workbook
from datetime import datetime
import requests

def capturar_clima():
    try:
        # Configurando a API
        api_key = "afe4b3825e114fb12c64e6f36d9ce369"  # Minha chave API
        cidade = "São Paulo"
        url = f"http://api.openweathermap.org/data/2.5/weather?q={cidade}&appid={api_key}&units=metric&lang=pt"
        
        # Fazendo a requisição à API
        response = requests.get(url)
        dados = response.json()

        # Extraindo informações relevantes
        temperatura = dados["main"]["temp"]
        umidade = dados["main"]["humidity"]
        descricao = dados["weather"][0]["description"]
        data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Abrindo ou criando a planilha
        try:
            workbook = load_workbook("dados_clima.xlsx")  # Tenta abrir a planilha existente
            sheet = workbook.active
        except FileNotFoundError:
            workbook = workbook()  # Cria uma nova planilha caso não exista
            sheet = workbook.active
            sheet.title = "Clima"
            sheet.append(["Data/Hora", "Temperatura (°C)", "Umidade (%)", "Descrição"])  # Adiciona cabeçalhos

        # Adicionando os dados na próxima linha
        sheet.append([data_hora, temperatura, umidade, descricao])
        
        # Salvando o arquivo
        workbook.save("dados_clima.xlsx")
        print("Dados adicionados com sucesso!")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Interface gráfica
root = tk.Tk()
root.geometry("150x100")
root.title("Capturar Clima")

button = tk.Button(root, text="🌞Capturar Clima☔", command=capturar_clima, font=("Arial", 12), padx=10, pady=5)
button.pack(pady=20)


root.mainloop()