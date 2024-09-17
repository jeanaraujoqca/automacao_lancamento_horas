import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import streamlit as st
from time import sleep
# import win32com.client as win32
import sys

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# # Função para enviar o relatório por email após a execução
# def enviar_relatorio():
#     outlook = win32.Dispatch('outlook.application')
#     namespace = outlook.GetNamespace('MAPI')

#     def achar_pasta_por_nome(nome_pasta, parent_folder=None):
#         if parent_folder is None:
#             parent_folder = namespace.Folders
            
#         for folder in parent_folder:
#             if folder.Name == nome_pasta:
#                 return folder
#             sub_folder = achar_pasta_por_nome(nome_pasta, folder.Folders)
#             if sub_folder:
#                 return sub_folder
#         return None 
    
#     sent_items_folder = achar_pasta_por_nome("Itens Enviados")
#     nome_remetente = 'Desconhecido' if not sent_items_folder else sent_items_folder.Items.GetLast().SenderName
    
#     mail = outlook.CreateItem(0)
#     mail.Subject = 'Relatório de Uso da Automação de Lançamento de Horas'
#     mail.Body = f'{nome_remetente} utilizou a automação de lançamento de horas.'
#     mail.To = 'daniellerodrigues@queirozcavalcanti.adv.br'
#     mail.Send()

# Função principal para executar a automação Playwright
async def submit_form(excel_file, email, senha):
    df = pd.read_excel(excel_file)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        # URL do SharePoint
        url_sharepoint = 'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'

        # Acessa o SharePoint
        await page.goto(url_sharepoint)
        await page.wait_for_timeout(5000)

        try:
            # Faz login
            await page.fill('#i0116', email)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.fill('#i0118', senha)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.click('#idSIButton9')  # Botão "Sim"
        except Exception as e:
            st.error(f'Erro ao fazer login: {str(e)}')
            return
        
        await page.wait_for_timeout(10000)
        st.write('Entramos no Sharepoint. Iniciando o preenchimento das informações.')

        casos_sucesso = []
        casos_fracasso = []

        # Itera pelo DataFrame e preenche os dados
        for index, id in enumerate(df['ID']):
            try:
                colaborador = df.loc[index, 'Nome']
                email_colaborador = df.loc[index, 'Email']
                unidade = df.loc[index, 'UNIDADE']
                treinamento = df.loc[index, 'TREINAMENTO']
                tipo_de_treinamento = df.loc[index, 'TIPO DO TREINAMENTO']
                categoria = df.loc[index, 'CATEGORIA']
                instituicao_instrutor = df.loc[index, 'INSTITUIÇÃO/INSTRUTOR']
                carga_horaria = df.loc[index, 'CARGA HORÁRIA']
                inicio_do_treinamento = df.loc[index, 'INICIO DO TREINAMENTO']
                termino_do_treinamento = df.loc[index, 'TERMINO DO TREINAMENTO']

                st.write(f'Adicionando informações do colaborador: {colaborador}')

                # Adiciona novo treinamento
                await page.click('button:has-text("Novo")')
                await page.wait_for_timeout(5000)
                
                # Muda para o iframe correto
                iframe = page.frame_locator("iframe").nth(0)
                iframe2 = iframe.frame_locator("iframe.player-app-frame")

                # Preenche as informações necessárias
                async def clica_seleciona_informacao(iframe, endereco1, endereco2, valor2, endereco3):
                    await iframe.locator(endereco1).click()  
                    await iframe.locator(endereco2).fill(valor2)
                    await iframe.locator(endereco3).nth(0).click()

                await clica_seleciona_informacao(iframe2, 'div[title="NOME DO INTEGRANTE"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input', colaborador, 
                                                 f'li:has-text("{colaborador}")')
                
                await clica_seleciona_informacao(iframe2, 'div[title="E-MAIL"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input', email_colaborador, 
                                                 f'li:has-text("{email_colaborador}")')
                
                await clica_seleciona_informacao(iframe2, 'div[title="UNIDADE"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input', unidade, 
                                                 f'li:has-text("{unidade}")')
                
                await iframe2.locator('input[title="TREINAMENTO"]').fill(treinamento)

                await clica_seleciona_informacao(iframe2, 'div[title="TIPO DO TREINAMENTO."]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input', tipo_de_treinamento, 
                                                 f'li:has-text("{tipo_de_treinamento}")')

                await iframe2.locator('input[title="INSTITUIÇÃO/INSTRUTOR"]').fill(instituicao_instrutor)
                await clica_seleciona_informacao(iframe2, 'div[title="CATEGORIA"]', 
                                                 '//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input', categoria, 
                                                 f'li:has-text("{categoria}")')
                
                await iframe2.locator('input[title="INICIO DO TREINAMENTO"]').fill(inicio_do_treinamento)
                await iframe2.locator('input[title="TERMINO DO TREINAMENTO"]').fill(termino_do_treinamento)

                # Salva os dados
                await page.locator('//*[@id="appRoot"]/div[3]/div/div[4]/div[2]/div/div[2]/div[3]/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[1]/button/span').click()
                
                st.write(f'{index+1} - ID {id} - {colaborador} - {treinamento} - finalizado')
                casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
                await asyncio.sleep(3)

            except Exception as e:
                st.error(f'Erro ao processar o treinamento {id}: {e}')
                casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

        # Gera os arquivos de sucesso e fracasso
        df_sucesso = pd.DataFrame(casos_sucesso)
        df_fracasso = pd.DataFrame(casos_fracasso)
        df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
        df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

        # # Envia o relatório por e-mail
        # enviar_relatorio()
        st.write('O processo finalizou! Verificar no Sharepoint as informações editadas.')
        await browser.close()

# Configura a interface do Streamlit
st.title('Automação de Lançamento de Horas de Treinamento')

# Entrada de dados
email = st.text_input('Email:')
senha = st.text_input('Senha:', type='password')
file = st.file_uploader('Escolha o arquivo Excel:', type=['xlsx'])

# Botão para iniciar a automação
if st.button('Iniciar Automação'):
    if email and senha and file:
        try:
            # Inicia a automação Playwright
            asyncio.run(submit_form(file, email, senha))
            st.success('Automação executada com sucesso!')
        except Exception as e:
            st.error(f'Erro ao executar a automação: {str(e)}')
    else:
        st.error('Por favor, preencha todos os campos!')
