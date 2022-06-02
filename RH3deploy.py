# -*- coding: utf-8 -*-
"""
Created on Tue May 31 20:23:59 2022

@author: marlo
"""
import pandas as pd
import numpy as np
import win32com.client as win32
from st_aggrid import AgGrid
import pandas as pd
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

st.title("Análise RH ")
st.subheader("Arquivo XLSX - Abaixo os Códigos dos Clientes com ou sem predição de default")
st.write('\n\n')

name = st.text_input('Coloque Seu Nome')
st.write('Nome:', name)

email = st.text_input('Coloque Seu E-mail')
st.write('E-mail:', email)

quant = st.number_input( 'Quanto tempo você tem na Organização?')
st.write('The current number is ', quant)

tempo_cargo_select = st.number_input( 'Quanto tempo você está no seu cargo atual?')
st.write('The current number is ', tempo_cargo_select)

cargo_atual_select = st.selectbox('Selecione qual seu cargo atual:',('Caixa','Gerente Assistente PF ou PJ','Gerente PJ','Gerente Exclusive','Gerente Comercial','Gerente Administrativo','Gerente Geral','Gerente de Pab','Supervisor Administrativo','Escriturário'))
st.write('You selected:', cargo_atual_select)


certifica = st.multiselect('Possui Alguma Certificação:',['CPA10','CPA20','CEA','CFP','CFA','Outras'])
st.write('You selected:', certifica)


cargo_almeja = st.selectbox('Qual próximo cargo que almeja:',['Gerente PJ','Gerente Geral','Gerente Exclusive','Gerente Assistente','Gerente Administrativo','Gerente Prime','Departamento','Nenhum Cargo'])
st.write('You selected:', cargo_almeja)

ponto_fraco_select = st.selectbox('Qual das competências abaixo você considera seu ponto fraco?',("Liderança","Planejamento estratégico","Criatividade","Relacionamento interpessoal","Equilíbrio emocional","Negociação","Visão globalizada","Percepção e julgamento","Empreendendorismo","Flexibilidade","Comunicação","Trabalho em equipe"))
st.write('You selected:', ponto_fraco_select)

ponto_forte_select = st.selectbox('Qual das competências abaixo você considera seu ponto forte',("Liderança","Planejamento estratégico","Criatividade","Relacionamento interpessoal","Equilíbrio emocional","Negociação","Visão globalizada","Percepção e julgamento","Empreendendorismo","Flexibilidade","Comunicação","Trabalho em equipe"))
st.write('You selected:', ponto_forte_select)

if st.button('Analisar'):
    livro_sugerido25 = str()
    curso_sugerido25 = str()
    curso_sugerido50 = str()
    livro_sugerido_geral = str()
    hardskill50 = str()
    if ponto_fraco_select == 'Comunicação':
        livro_sugerido25 = 'Apremda a se comunicar com habilidade e clareza. Autor: Arrendondo, Lani'
        curso_sugerido25 = 'Na trilha da comunicação eficaz'
        curso_sugerido50 = 'Técnicas de Apresentação'
    if ponto_fraco_select == 'Criatividade':
        livro_sugerido25 = 'Um TOC na CUCA. Autor: Von Oech, Roger'
        curso_sugerido25 = 'Inovação'
    if cargo_almeja == 'Gerente PJ':
        livro_sugerido_geral = 'A Arte de Argumentar - Gerenciando Raz"ao e Emoção.Autor: Antonio Suarez'
        curso_sugerido_geral = 'Teoria e Prática na Negociação'
        desempenho25 = 'Leitura Manual Pade'
        desempenho50 ='Reunião detalhamento orçado e realizado'
        desempenho75 = 'Apresentação Pade aos colegas'
        hardskill50 = 'Curso Matemática Financeira'
    if certifica == 'CPA10':
        certifica25 = 'Inscrição Curso Integra a realizar em até 2 meses'
        certifica50 = 'Inscrição para Prova de 2 a 3 meses após estudo'
    dfgeral = pd.DataFrame({'Competência 1': [],'Competência 2': [],'HardSkill': [],'Desempenho':[],'Certificações':[],'Prazo':[]})
    dfgeral.loc[0, 'Competência 1'] = livro_sugerido25
    dfgeral.loc[0, 'Competencia 2'] = livro_sugerido_geral
    dfgeral.loc[1, 'HardSkill'] = hardskill50
    AgGrid(dfgeral)
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    df_xlsx = to_excel(dfgeral)
    st.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'df_test.xlsx')




    #st.success(dfgeral)
#dfgeral.to_excel(r'D:\MARLON\Teste_PDI.xlsx', sheet_name='Your sheet name', index = False)
#if st.button('Enviar e-mail'):
#    outlook = win32.Dispatch('outlook.application')
    # criar um email
#    email = outlook.CreateItem(0)
# configurar as informações do seu e-mail
 #   email.To = "marlon.engamb@gmail.com"
   # email.Subject = "E-mail automático do Python"
  #  email.HTMLBody = f"""
   # <p>Olá Lira, aqui é o código Python</p>
   # """
   # anexo = dfgeral
    # email.Attachments.Add(anexo)
   # email.Send()
   # st.success('E-mail Enviado')

    





