import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pywhatkit as kit
import time
from datetime import datetime, timedelta

# Dados dos clientes
clientes = [
    {"Nome": "Rafael", "Numero": "+5541999630682"},
    {"Nome": "Pastora Cris", "Numero": "+5541987007055"},
    # Adicione mais clientes conforme necessário
]

# CONFIGURAÇÕES PRINCIPAIS (MODIFIQUE AQUI)
hora_inicio = 6  # Hora desejada (formato 24h)
minuto_inicio = 31  # Minuto desejado
intervalo_segundos = 120  # 120 segundos = 2 minutos entre mensagens

# Configurações do Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\vscode\google\dev-exchanger-437322-m7-7de286a5b6c1.json', scope)
client = gspread.authorize(creds)

# Abrir a planilha
sheet_id = "1lkzwzlmfHXEda4cxEwo13zmZwUd1klpGSK06Vi5-QQw"
sheet = client.open_by_key(sheet_id)

# Função para buscar dados
def buscar_dados():
    worksheet = sheet.worksheet("VALOR")
    return worksheet.get("A2:B14"), worksheet.get("A16:B21")

# Função para montar mensagem
def criar_mensagem(nome, producao, para_produzir):
    msg = f"Olá, bom dia {nome}! 🌞\n\nItens para produção:\n"
    msg += "\n".join([f"{i[0]} -> {i[1]}" for i in producao if i[0] and i[1]])
    msg += "\n\nPara produzir:\n"
    msg += "\n".join([f"{i[0]} -> {i[1]}" for i in para_produzir if len(i) >= 2 and i[0] and i[1]])
    return msg + "\n\nObrigado, que Deus abençoe hoje, e seja um ótimo dia! 🙌🙏😊"

# Verificar e ajustar horário
agora = datetime.now()
inicio = agora.replace(hour=hora_inicio, minute=minuto_inicio, second=0, microsecond=0)

# Se horário já passou, começa IMEDIATAMENTE com intervalos
if agora > inicio:
    inicio = agora + timedelta(seconds=30)  # Começa em 30 segundos
    print(f"⚠️ Horário inicial já passou, iniciando em: {inicio.strftime('%H:%M:%S')}")

# Buscar dados
dados_producao, dados_para_produzir = buscar_dados()

# ENVIO DAS MENSAGENS
print("\n📅 AGENDA DE ENVIOS 📅")
for i, cliente in enumerate(clientes):
    if not cliente['Numero'] or not cliente['Numero'].strip():
        print(f"❌ {cliente['Nome']}: Número inválido")
        continue
    
    horario_envio = inicio + timedelta(seconds=i * intervalo_segundos)
    mensagem = criar_mensagem(cliente['Nome'], dados_producao, dados_para_produzir)
    
    print(f"\n🔔 {i+1}. {cliente['Nome']}")
    print(f"⏰ Horário: {horario_envio.strftime('%H:%M:%S')}")
    print(f"📱 Número: {cliente['Numero']}")
    
    try:
        kit.sendwhatmsg(
            phone_no=cliente['Numero'].strip(),
            message=mensagem,
            time_hour=horario_envio.hour,
            time_min=horario_envio.minute,
            wait_time=15,
            tab_close=True
        )
        print("✅ AGENDADO COM SUCESSO!")
    except Exception as e:
        print(f"❌ ERRO: {str(e)}")
    
    if i < len(clientes) - 1:
        time.sleep(4)  # Pequena pausa entre agendamentos

print("\n🎉 TODAS AS MENSAGENS FORAM AGENDADAS PARA HOJE!")
print("⏳ Horários dos envios:")
for i, cliente in enumerate(clientes):
    if cliente['Numero'].strip():
        print(f"{i+1}. {cliente['Nome']}: {inicio + timedelta(seconds=i * intervalo_segundos)}")