import os
import csv
import json
import mailbox
from icalendar import Calendar

def converter_ics_para_csv(caminho_arquivo, pasta_saida):
    try:
        with open(caminho_arquivo, 'rb') as f:
            cal = Calendar.from_ical(f.read())

        eventos = []
        for componente in cal.walk():
            if componente.name == "VEVENT":
                evento = {
                    "Resumo": str(componente.get("summary")),
                    "Início": str(componente.get("dtstart").dt),
                    "Fim": str(componente.get("dtend").dt),
                    "Descrição": str(componente.get("description")),
                    "Local": str(componente.get("location"))
                }
                eventos.append(evento)

        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".csv"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        with open(caminho_saida, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=eventos[0].keys())
            writer.writeheader()
            writer.writerows(eventos)

        return True, caminho_saida
    except Exception as e:
        return False, str(e)

def converter_ics_para_json(caminho_arquivo, pasta_saida):
    try:
        with open(caminho_arquivo, 'rb') as f:
            cal = Calendar.from_ical(f.read())

        eventos = []
        for componente in cal.walk():
            if componente.name == "VEVENT":
                evento = {
                    "Resumo": str(componente.get("summary")),
                    "Início": str(componente.get("dtstart").dt),
                    "Fim": str(componente.get("dtend").dt),
                    "Descrição": str(componente.get("description")),
                    "Local": str(componente.get("location"))
                }
                eventos.append(evento)

        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".json"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        with open(caminho_saida, 'w', encoding='utf-8') as f:
            json.dump(eventos, f, ensure_ascii=False, indent=4)

        return True, caminho_saida
    except Exception as e:
        return False, str(e)

def converter_mbox_para_csv(caminho_arquivo, pasta_saida):
    try:
        mbox = mailbox.mbox(caminho_arquivo)
        mensagens = []

        for msg in mbox:
            mensagens.append({
                "De": msg.get("From"),
                "Para": msg.get("To"),
                "Assunto": msg.get("Subject"),
                "Data": msg.get("Date")
            })

        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".csv"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        with open(caminho_saida, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=mensagens[0].keys())
            writer.writeheader()
            writer.writerows(mensagens)

        return True, caminho_saida
    except Exception as e:
        return False, str(e)

def converter_mbox_para_json(caminho_arquivo, pasta_saida):
    try:
        mbox = mailbox.mbox(caminho_arquivo)
        mensagens = []

        for msg in mbox:
            mensagens.append({
                "De": msg.get("From"),
                "Para": msg.get("To"),
                "Assunto": msg.get("Subject"),
                "Data": msg.get("Date")
            })

        nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".json"
        caminho_saida = os.path.join(pasta_saida, nome_arquivo)

        with open(caminho_saida, 'w', encoding='utf-8') as f:
            json.dump(mensagens, f, ensure_ascii=False, indent=4)

        return True, caminho_saida
    except Exception as e:
        return False, str(e)
