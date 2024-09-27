import streamlit as st
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import os

async def submit_form(email, senha, file):
    # Read the uploaded Excel file
    df = pd.read_excel(file)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # SharePoint URL
        url_sharepoint = 'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'
        await page.goto(url_sharepoint)
        await page.wait_for_timeout(5000)

        try:
            # Realiza login
            await page.fill('#i0116', email)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.fill('#i0118', senha)
            await page.click('#idSIButton9')
            await page.wait_for_timeout(2000)
            await page.click('#idSIButton9')  # Botão "Sim" para manter logado
        except Exception as e:
            st.error(f"Erro no login: {e}")
            return

        await page.wait_for_timeout(10000)
        st.write('Entramos no Sharepoint. Aguarde para iniciar o procedimento de preenchimento das informações dos treinamentos.')

        # Process the rows from the Excel file
        casos_sucesso = []
        casos_fracasso = []

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

                # Add new training information
                await page.click('button:has-text("Novo")')
                await page.wait_for_timeout(30000)

                # Switch to the correct iframe
                iframe = page.frame_locator("iframe").nth(0)
                iframe2 = iframe.frame_locator("iframe.player-app-frame")

                # Function to interact with elements
                async def clica_seleciona_informacao(iframe, endereco1, endereco2, valor2, endereco3):
                    await iframe.locator(endereco1).click()
                    await iframe.locator(endereco2).fill(valor2)
                    await page.wait_for_timeout(1000)
                    await iframe.locator(endereco3).nth(0).click()

                # Fill the data from the Excel sheet
                await clica_seleciona_informacao(iframe2, 'div[title="NOME DO INTEGRANTE"]',
                                                 '//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input', colaborador,
                                                 f'li:has-text("{colaborador}")')
                await clica_seleciona_informacao(iframe2, 'div[title="E-MAIL"]',
                                                 '//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input', email_colaborador,
                                                 f'li:has-text("{email_colaborador}")')
                await clica_seleciona_informacao(iframe2, 'div[title="UNIDADE"]',
                                                 '//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input', unidade,
                                                 f'li:has-text("{unidade}")')

                # Fill other fields
                await iframe2.locator('input[title="TREINAMENTO"]').fill(treinamento)
                await clica_seleciona_informacao(iframe2, 'div[title="TIPO DO TREINAMENTO."]',
                                                 '//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input', tipo_de_treinamento,
                                                 f'li:has-text("{tipo_de_treinamento}")')
                await clica_seleciona_informacao(iframe2, 'div[title="CATEGORIA"]',
                                                 '//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input', categoria,
                                                 f'li:has-text("{categoria}")')
                await iframe2.locator('input[title="INSTITUIÇÃO/INSTRUTOR"]').fill(instituicao_instrutor)
                await iframe2.locator('input[title="INICIO DO TREINAMENTO"]').fill(inicio_do_treinamento)
                await iframe2.locator('input[title="TERMINO DO TREINAMENTO"]').fill(termino_do_treinamento)

                # Save the data
                await page.locator('//*[@id="appRoot"]/div[3]/div/div[4]/div[2]/div/div[2]/div[3]/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[1]/button/span').click()

                st.write(f'{index+1} - ID {id} - {colaborador} - {treinamento} - finalizado')
                casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
                await asyncio.sleep(3)

            except Exception as e:
                st.error(f'Erro inesperado ao processar o treinamento {id}: {e}')
                casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

        # Save the results
        df_sucesso = pd.DataFrame(casos_sucesso)
        df_fracasso = pd.DataFrame(casos_fracasso)
        df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
        df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

        await browser.close()
        st.write('O processo finalizou! Verificar no Sharepoint as informações editadas.')

# Streamlit app interface
st.title('Automação de Lançamento de Horas de Treinamento')

email = st.text_input('Email:')
senha = st.text_input('Senha:', type='password')
file = st.file_uploader('Escolha o arquivo Excel:', type=['xlsx'])

if st.button('Iniciar Automação'):
    if email and senha and file:
        asyncio.run(submit_form(email, senha, file))
    else:
        st.error('Por favor, preencha todos os campos!')
