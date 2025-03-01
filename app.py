# Importações necessárias
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Função para enviar mensagem e imagem via WhatsApp Web
def enviar_mensagem_whatsapp(nome, telefone, mensagem, caminho_imagem=None):
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)  # Tempo para o WhatsApp Web carregar

        # Se uma imagem for fornecida, anexe e envie
        if caminho_imagem:
            try:
                # Localizar o botão de anexo (clipe)
                botao_anexo = pyautogui.locateCenterOnScreen('Selecio.png', confidence=0.8)
                if botao_anexo:
                    pyautogui.click(botao_anexo)  # Clicar no botão de anexo
                    sleep(2)

                    # Localizar e clicar no botão "Foto e Vídeo"
                    botao_foto_video = pyautogui.locateCenterOnScreen('FotoVi.png', confidence=0.8)
                    if botao_foto_video:
                        pyautogui.click(botao_foto_video)
                        sleep(2)

                        # Digitar o caminho da imagem e pressionar Enter
                        pyautogui.write(caminho_imagem)
                        sleep(1)
                        pyautogui.press('enter')
                        sleep(5)  # Tempo para a imagem ser carregada
                    else:
                        print("Botão 'Foto e Vídeo' não encontrado.")
                else:
                    print("Botão de anexo não encontrado.")
            except Exception as e:
                print(f"Erro ao anexar a imagem: {e}")

        # Enviar a mensagem (e a imagem, se anexada)
        try:
            seta = pyautogui.locateCenterOnScreen('seta.png', confidence=0.8)
            if seta:
                print(f"Setinha encontrada na posição: {seta}")
                pyautogui.click(seta[0], seta[1])
            else:
                print("Seta não encontrada. Tentando enviar com a tecla Enter...")
                pyautogui.press('enter')

            sleep(2)
            pyautogui.hotkey('ctrl', 'w')  # Fechar a aba
            print(f"Mensagem e imagem enviadas para {nome} ({telefone}).")
        except Exception as e:
            print(f"Erro ao localizar a seta ou enviar a mensagem: {e}")
    except Exception as e:
        print(f"Erro ao enviar mensagem para {nome}: {e}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')

# Abrir WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)  # Tempo para o usuário escanear o QR Code

# Ler planilha e enviar mensagens  
try:
    workbook = openpyxl.load_workbook('Clientes.xlsx')
    pagina_clientes = workbook['Planilha1']

    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = getattr(linha[0], 'value', None)  # Nome
        telefone = getattr(linha[1], 'value', None)  # Telefone

        # Verifica se todos os campos obrigatórios estão preenchidos
        if not nome or not telefone:
            print(f"Dados incompletos na linha {linha[0].row}. Pulando para a próxima.")
            continue

        mensagem = f"Olá {nome}! 👋, Você conhece os nossos produtos da Labtest? O Betatest é um imunoensaio cromatográfico desenvolvido pela Labtest Diagnóstica para a detecção qualitativa rápida da Gonadotrofina Coriônica Humana (hCG) em amostras de soro ou urina. Este teste é amplamente utilizado para a confirmação precoce de gravidez.Princípio do Teste:Durante o procedimento, a amostra é colocada em contato com um conjugado de ouro coloidal ligado a anticorpos monoclonais anti-hCG. Se a amostra contiver hCG, formará um complexo que migrará por ação capilar através da membrana do dispositivo. Ao alcançar a região onde anticorpos anti-hCG estão imobilizados, uma linha colorida aparecerá, indicando um resultado positivo. Uma segunda linha de controle, que deve sempre aparecer, confirma a validade e a correta execução do teste."
        caminho_imagem = "BETATEST.png"  # Caminho da imagem
        print(f"Processando contato: {nome} ({telefone})...")
        enviar_mensagem_whatsapp(nome, telefone, mensagem, caminho_imagem)

except Exception as e:
    print(f"Erro ao processar a planilha: {e}")