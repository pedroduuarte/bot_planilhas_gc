import win32com.client
import time
import telebot
from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup
import pythoncom
import os
from dotenv import load_dotenv


BASE_DIR = r"C:\Users\duart\OneDrive\Área de Trabalho\Codes\bot-planilhas-alfa\dados_ligas"

ligas = {
    "Premier League": "PREMIER.xlsx",
    "La Liga": "LA LIGA.xlsx",
    "Bundesliga": "BUNDES.xlsx",
    "Serie A": "SERIE A.xlsx"
}

load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
bot = telebot.TeleBot(TOKEN)

user_data = {}

# Função para analisar qualquer mercado
def analisar_mercados(ws, linha_odds, linha_valores, col_inicio, col_fim, descricao):
    odd_referencia = 1.3
    odds_dados = [(ws.Cells(linha_odds, col).Value, ws.Cells(linha_valores, col).Value) for col in range(col_inicio, col_fim + 1)]

    odds_dados = [
        (round(odd, 2), valor) for odd, valor in odds_dados if isinstance(odd, (int, float)) and isinstance(valor, (int, float))
    ]

    if not odds_dados:
        return f"Nenhuma odd válida encontrada no mercado de {descricao}"
        

    # Encontra a odd mais próxima abaixo e acima de 1.3
    odd_mais_proxima = min(
        odds_dados, key=lambda x: (abs(x[0] - odd_referencia), x[0] > odd_referencia)
    )

    return f"📈 {descricao}: {odd_mais_proxima[1]:.1f} (Odd: {odd_mais_proxima[0]:.2f})"

@bot.message_handler(commands=['start'])
def iniciar_conversa(message):
    markup = InlineKeyboardMarkup()
    for nome_liga, arq in ligas.items():
        markup.add(InlineKeyboardButton(text=nome_liga, callback_data=arq))

    bot.send_message(message.chat.id, "Escolha a liga que você deseja analisar:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data in ligas.values())
def receber_liga(call):
    user_data[call.message.chat.id] = {"liga": call.data}
    bot.send_message(call.message.chat.id, "Agora informe o time da casa:")
    


@bot.message_handler(func=lambda message: message.chat.id in user_data and "time_casa" not in user_data[message.chat.id])
def receber_time_casa(message):
    user_data[message.chat.id]["time_casa"] = message.text.upper()
    bot.send_message(message.chat.id, "Agora, informe o time visitante:")

@bot.message_handler(func=lambda message: message.chat.id in user_data and "time_visitante" not in user_data[message.chat.id])
def receber_time_visitante(message):
    user_data[message.chat.id]["time_visitante"] = message.text.upper()
    bot.send_message(message.chat.id, "Por fim, informe o árbitro do jogo:")

@bot.message_handler(func=lambda message: message.chat.id in user_data and "arbitro" not in user_data[message.chat.id])
def receber_arbitro(message):
    user_data[message.chat.id]["arbitro"] = message.text.upper()
    bot.send_message(message.chat.id, "Processando os dados... Aguarde alguns segundos.")

    # Inicia o Excel via COM
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  

    # obtém o caminho do arquivo da liga escolhida
    arquivo_excel = os.path.join(BASE_DIR, user_data[message.chat.id]["liga"])

    try:
        # Abre a planilha
        wb_excel = excel.Workbooks.Open(arquivo_excel)
    except Exception as e:
        bot.send_message(message.chat.id, f"Erro ao abrir o arquivo: {e}")
        excel.Quit()
        return
    
    ws = wb_excel.Sheets("Simulação Jogos")

    # Preenche a planilha com os dados informados
    ws.Range("A2").Value = user_data[message.chat.id]["time_casa"]
    ws.Range("C2").Value = user_data[message.chat.id]["time_visitante"]
    ws.Range("A8").Value = user_data[message.chat.id]["arbitro"]

    # Força o recálculo das fórmulas
    excel.Application.CalculateFullRebuild()  # Recalcula todas as células
    # Espera para garantir que o cálculo foi feito
    time.sleep(3)

    # Analisar os mercados OVER
    mensagem = "Análise de Mercados OVER\n\n"
    mensagem += analisar_mercados(ws, 4, 5, 55, 78, "Faltas Totais") + "\n"
    mensagem += analisar_mercados(ws, 8, 9, 55, 78, "Faltas CASA") + "\n"
    mensagem += analisar_mercados(ws, 10, 11, 55, 78, "Faltas VISITANTE") + "\n"
    mensagem += analisar_mercados(ws, 14, 15, 55, 78, "Finalizações Total") + "\n"
    mensagem += analisar_mercados(ws, 18, 19, 55, 78, "Finalizações CASA") + "\n"
    mensagem += analisar_mercados(ws, 20, 21, 55, 78, "Finalizações VISITANTE") + "\n"
    mensagem += analisar_mercados(ws, 24, 25, 55, 78, "Chutes ao gol Total") + "\n"
    mensagem += analisar_mercados(ws, 28, 29, 55, 78, "Chute ao Gol CASA") + "\n"
    mensagem += analisar_mercados(ws, 30, 31, 55, 78, "Chute ao Gol VISITANTE") + "\n"
    mensagem += analisar_mercados(ws, 34, 35, 55, 72, "Cartões Totais") + "\n"
    mensagem += analisar_mercados(ws, 38, 39, 55, 67, "Cartões CASA") + "\n"
    mensagem += analisar_mercados(ws, 40, 41, 55, 67, "Cartões VISITANTE") + "\n"
    mensagem += analisar_mercados(ws, 44, 45, 55, 78, "Desarmes Totais") + "\n"
    mensagem += analisar_mercados(ws, 48, 49, 55, 78, "Desarmes CASA") + "\n"
    mensagem += analisar_mercados(ws, 50, 51, 55, 78, "Desarmes VISITANTE") + "\n"

    # Analisar os mercados UNDER
    mensagem2 = "Análise de Mercados UNDER\n\n"
    mensagem2 += analisar_mercados(ws, 4, 5, 81, 102, "Faltas Totais") + "\n"
    mensagem2 += analisar_mercados(ws, 8, 9, 81, 102, "Faltas CASA") + "\n"
    mensagem2 += analisar_mercados(ws, 10, 11, 81, 102, "Faltas VISITANTE")+ "\n"
    mensagem2 += analisar_mercados(ws, 14, 15, 81, 102, "Finalizações Total") + "\n"
    mensagem2 += analisar_mercados(ws, 18, 19, 81, 102, "Finalizações CASA") + "\n"
    mensagem2 += analisar_mercados(ws, 20, 21, 81, 102, "Finalizações VISITANTE") + "\n"
    mensagem2 += analisar_mercados(ws, 24, 25, 81, 102, "Chutes ao gol total") + "\n"
    mensagem2 += analisar_mercados(ws, 28, 29, 81, 102, "Chute ao Gol CASA") + "\n"
    mensagem2 += analisar_mercados(ws, 30, 31, 81, 102, "Chute ao Gol VISITANTE") + "\n"
    mensagem2 += analisar_mercados(ws, 34, 34, 81, 98, "Cartões Totais") + "\n"
    mensagem2 += analisar_mercados(ws, 38, 39, 81, 93, "Cartões CASA") + "\n"
    mensagem2 += analisar_mercados(ws, 40, 41, 81, 93, "Cartões VISITANTE") + "\n"
    mensagem2 += analisar_mercados(ws, 44, 45, 81, 102, "Desarmes Totais") + "\n"
    mensagem2 += analisar_mercados(ws, 48, 49, 81, 102, "Desarmes CASA") + "\n"
    mensagem2 += analisar_mercados(ws, 50, 51, 81, 102, "Desarmes VISITANTE") + "\n"

    # Salva e fecha o arquivo
    wb_excel.Save()
    wb_excel.Close()
    excel.Quit()

    # Envia os resultados pelo Telegram
    bot.send_message(message.chat.id, mensagem)
    bot.send_message(message.chat.id, mensagem2)

bot.infinity_polling(timeout=60, long_polling_timeout=10)
