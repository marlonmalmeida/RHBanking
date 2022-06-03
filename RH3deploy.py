# -*- coding: utf-8 -*-
"""
Created on Tue May 31 20:23:59 2022

@author: marlo
"""
import pandas as pd
import numpy as np
#import win32com.client as win32
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
        livro_sugerido50 = 'Leitura do livro: Aprenda a se comunicar com habilidade e clareza. Autor: Arrendondo, Lani'
        curso_sugerido25 = 'Curso : Na trilha da comunicação eficaz'
        filme_sugerido50 = 'Filme sugerido: O informante'
    if ponto_fraco_select == 'Criatividade':
        livro_sugerido50 = 'Leitura do Livro: Um TOC na CUCA. Autor: Von Oech, Roger'
        curso_sugerido25 = 'Curso: Inovação'
        filme_sugerido50 = 'Filme: Vem Dançar'
    if ponto_fraco_select == 'Empreendedorismo':
        livro_sugerido50 = 'Leitura do Livro: Nos bastidores da Disney.Autor: Tom Connellan'
        curso_sugerido25 = 'Curso: Sem indicações'
        filme_sugerido50 = 'Filme: À procura da felicidade'
    if ponto_fraco_select == 'Equilíbrio emocional':
        livro_sugerido50 = 'Leitura do Livro: O Caçador de Pipas. Autor: Khaled Hosseini'
        curso_sugerido25 = 'Curso: Sem indicações'
        filme_sugerido50 = 'Filme: Homens de Honra'
    if ponto_fraco_select == 'Flexibilidade':
        livro_sugerido50 = 'Leitura do Livro: Caminhos da mudança- Reflexões sobre um mundo impermanten...Autor: Eugênio Mussak'
        curso_sugerido25 = 'Curso: Sem indicações'
        filme_sugerido50 = 'Filme: O diabo veste Prada'
    if ponto_fraco_select == 'Liderança':
        livro_sugerido50 = 'Leitura do Livro: Você não precisa ser chefe para ser Líder. Autor: Mark Sanborn'
        curso_sugerido25 = 'Curso: A aventura da Liderança'
        filme_sugerido50 = 'Filme: Coração Valente'
    if ponto_fraco_select == 'Negociação':
        livro_sugerido50 = 'Leitura do Livro: A arte de Argumentar. Autor: Antonio Suarez Abreu'
        curso_sugerido25 = 'Curso: Ted - A arte da negociação'
        filme_sugerido50 = 'Filme: O Articulador'
    if ponto_fraco_select == 'Percepção e Julgamento':
        livro_sugerido50 = 'Leitura do Livro: Sem indicação'
        curso_sugerido25 = 'Curso: Sem indicação'
        filme_sugerido50 = 'Filme: Prenda-me se For Capaz'
    if ponto_fraco_select == 'Planejamento Estratégico':
        livro_sugerido50 = 'Leitura do Livro: A arte da Estratégia. Autor: Carlos alberto Julio'
        curso_sugerido25 = 'Curso: Sem sugestão'
        filme_sugerido50 = 'Filme: Os infiltrados'
    if ponto_fraco_select == 'Relacionamento Interpessoal':
        livro_sugerido50 = 'Leitura do Livro: Livre'
        curso_sugerido25 = 'Curso: O segredo da neurociência no relacionamento'
        filme_sugerido50 = 'Filme: Patch Adamns - O amor é contagioso'
    if ponto_fraco_select == 'Trabalho em Equipe':
        livro_sugerido50 = 'Leitura do Livro: Livre'
        curso_sugerido25 = 'Curso: Workshop Empatia'
        filme_sugerido50 = 'Filme: Onze homens e um segredo'
    if ponto_fraco_select == 'Visão Globalizada':
        livro_sugerido50 = 'Leitura do Livro: Livre'
        curso_sugerido25 = 'Curso: Zoom: Visão globalizada'
        filme_sugerido50 = 'Filme: Wall Street: O dinheiro nunca dorme'
    if cargo_almeja == 'Gerente PJ':
        livro_sugerido_geral = 'A Arte de Argumentar - Gerenciando Razão e Emoção.Autor: Antonio Suarez'
        curso_sugerido_geral = 'Teoria e Prática na Negociação'
        hab75_geral = 'Estágio 1 Semana Gerente PJ'
        hab75 = 'Delegá-lo Acompanhar Produção de um Produto'
        hab_100_geral = 'Negociar com um cliente uma linha de crédito com Supervisão'
        hab_100 = 'Montar Planejamento da sua semana e apresentar ao Gerente.'
        desempenho25 = 'Leitura Manual Pade'
        desempenho50 ='Reunião detalhamento orçado e realizado'
        desempenho75 = 'Apresentação Pade aos colegas'
        desempenho100 = 'Apresentar orçado e realizado com destaque aos principais produtos'	
        hardskill50 = 'Curso Matemática Financeira'
        hardskill75 = 'Visita a um cliente Alto Valor'
        hardskill100 = 'Curso sobre Finanças Pessoais'
    if cargo_almeja == 'Gerente Exclusive':
        livro_sugerido_geral = 'O caçador de Pipas. Autor: Khaled Hosseini'
        curso_sugerido_geral = 'Curso: Aprenda a usar o Stress a seu favor'
        hab75_geral = 'Estágio 1 Semana Gerente Exclusive'
        hab75 = 'Delegá-lo Acompanhar Produção de um Produto'
        hab_100_geral = 'Negociar com um cliente uma linha de crédito com Supervisão'
        hab_100 = 'Montar Planejamento da sua semana e apresentar ao Gerente.'
        desempenho25 = 'Leitura Manual Pade'
        desempenho50 ='Reunião detalhamento orçado e realizado'
        desempenho75 = 'Apresentação Pade aos colegas'
        desempenho100 = 'Apresentar orçado e realizado com destaque aos principais produtos'	
        hardskill50 = 'Curso Matemática Financeira'
        hardskill75 = 'Visita a um cliente' 
        hardskill100 = 'Curso sobre Finanças Pessoais'
    if cargo_almeja == 'Gerente Prime':
        livro_sugerido_geral = 'O caçador de Pipas. Autor: Khaled Hosseini'
        curso_sugerido_geral = 'Curso: Aprenda a usar o Stress a seu favor'
        hab75_geral = 'Estágio 1 Semana Gerente Prime'
        hab75 = 'Delegá-lo Acompanhar Produção de um Produto'
        hab_100_geral = 'Negociar com um cliente uma linha de crédito com Supervisão'
        hab_100 = 'Montar Planejamento da sua semana e apresentar ao Gerente.'
        desempenho25 = 'Leitura Manual Pade'
        desempenho50 ='Reunião detalhamento orçado e realizado'
        desempenho75 = 'Apresentação Pade aos colegas'
        desempenho100 = 'Apresentar orçado e realizado com destaque aos principais produtos'	
        hardskill50 = 'Curso Matemática Financeira'
        hardskill75 = 'Visita a um cliente '
        hardskill100 = 'Curso sobre Finanças Pessoais'
    if cargo_almeja == 'Gerente Geral':
        livro_sugerido_geral = 'A geração Y no trabalho. Autor: Nicole Lipkin'
        curso_sugerido_geral = 'Curso: A arte de dar Feedback'
        hab75_geral = 'Estágio 1 Mês como Imediato'
        hab75 = 'Delegá-lo Acompanhar Produção da equipe'
        hab_100_geral = 'Fazer reunião com equipe sob supervisão'
        hab_100 = 'Montar Planejamento estratégico da agência.'
        desempenho25 = 'Leitura Manual POBJ'
        desempenho50 ='Estar bem qualificado no AGP'
        desempenho75 = 'Apresentação POBJ aos colegas - Resumo'
        desempenho100 = 'Apresentar orçado e realizado detalhadamente(Se possível Curso de Orçamento)'	
        hardskill50 = 'Curso Liderança'
        hardskill75 = 'Reunião Prestação de Contas com Regional '
        hardskill100 = 'Curso sobre Gestão de Pessoas'
    if cargo_almeja == 'Departamento':
        livro_sugerido_geral = 'A geração Y no trabalho. Autor: Nicole Lipkin'
        curso_sugerido_geral = 'Curso: Gestão da Mudança'
        hab75_geral = 'Contactar Gestor Área para bate papo sobre cargo desejado'
        hab75 = 'Habilidade Técnica 1 Exigida para Cargo'
        hab_100_geral = 'Fazer Checklist das habilidades adquiridas'
        hab_100 = 'Montar Planejamento estratégico de mudança de área.'
        desempenho25 = 'Leitura Manual POBJ/PADE'
        desempenho50 ='Estar bem qualificado no AGP'
        desempenho75 = 'Apresentação POBJ/PADE aos colegas - Resumo'
        desempenho100 = 'Apresentar orçado e realizado detalhadamente(Se possível Curso de Orçamento)'	
        hardskill50 = 'Verificar Habilidades Exigidas para Cargo como: Certificação, Cursos Extras, Formação Acadêmica'
        hardskill75 = 'Curso Treinet Gestão da Mudança '
        hardskill100 = 'Conclusão dos Cursos Especificos exigidos'
    if cargo_almeja == 'Gerente Administrativo':
        livro_sugerido_geral = 'Estratégia do Oceano Azul. Autor: W.Chan'
        curso_sugerido_geral = 'Curso: Gestão de Tempo'
        hab75_geral = 'Estágio 1 Semana na área adminstrativa'
        hab75 = 'Delegá-lo Acompanhar Gerenciamento da Equipe por 1 mes'
        hab_100_geral = 'Fazer reunião com equipe sob supervisão'
        hab_100 = 'Montar Planejamento estratégico da agência.'
        desempenho25 = 'Leitura Manual POBJ'
        desempenho50 ='Estar bem qualificado no AGP'
        desempenho75 = 'Apresentação POBJ aos colegas - Resumo área Adminstrativa'
        desempenho100 = 'Apresentar orçado e realizado detalhadamente(Se possível Curso de Orçamento)'	
        hardskill50 = 'Curso Liderança'
        hardskill75 = 'Aprender sobre diversas contas e razões contábeis '
        hardskill100 = 'Curso sobre Gestão de Pessoas'
    if certifica == ['CPA10']:
        certifica25 = 'Inscrição Curso CPA20 Integra a realizar em até 2 meses'
        certifica50 = 'Inscrição para Prova de 2 a 3 meses após estudo'
        certifica100 = 'Fazer escola sobre o que aprendeu com certificação para os demais colegas'
    if certifica == ['CPA10','CPA20']:
        certifica25 = 'Inscrição Curso CEA a realizar em até 2 meses'
        certifica50 = 'Inscrição para Prova'
        certifica100 = 'Fazer escola sobre o que aprendeu com certificação para os demais colegas'
    if certifica == ['CPA10','CPA20','CEA']:
        certifica25 = 'Inscrição Curso CFP'
        certifica50 = 'Inscrição para Prova de 6 a 8 meses após estudo'
        certifica100 = 'Se preparar para CFA'
    if certifica == ['CPA10','CPA20','CEA','CFP']:
        certifica25 = 'Inscrição Curso CFA primeiro modulo até 8 meses após inicio estudo'
        certifica50 = 'Inscrição para 2 Módulo da Prova de 12 a 24 meses após estudo'
        certifica100 = 'Incrição para 3 Módulo da Prova ate 36 meses'
    prazo1 = 'até 2 meses'
    prazo2 = 'até 3 meses'
    prazo3 = 'até 6 meses'
    prazo4 = 'até 12 meses'
    dfgeral = pd.DataFrame({'Cronograma':[],'Competência Base': [],'Competência Menor Desempenho': [],'HardSkill': [],'Desempenho':[],'Certificações':[]})
    dfgeral.loc[0, 'Competência Base'] = curso_sugerido25
    dfgeral.loc[0, 'Competência Menor Desempenho'] = curso_sugerido_geral
    dfgeral.loc[0, 'HardSkill'] = livro_sugerido_geral
    dfgeral.loc[0, 'Desempenho'] = desempenho25
    dfgeral.loc[0, 'Certificações'] = certifica25
    dfgeral.loc[0, 'Cronograma'] = prazo1
    dfgeral.loc[1, 'Competência Base'] = livro_sugerido50
    dfgeral.loc[1, 'Competência Menor Desempenho'] = str('-')
    dfgeral.loc[1, 'HardSkill'] = hardskill50
    dfgeral.loc[1, 'Desempenho'] = desempenho50
    dfgeral.loc[1, 'Certificações'] = certifica50
    dfgeral.loc[1, 'Cronograma'] = prazo2
    dfgeral.loc[2, 'Competência Base'] = hab75
    dfgeral.loc[2, 'Competência Menor Desempenho'] = hab75_geral
    dfgeral.loc[2, 'HardSkill'] = hardskill50
    dfgeral.loc[2, 'Desempenho'] = hardskill75
    dfgeral.loc[2, 'Certificações'] = str('-')
    dfgeral.loc[2, 'Cronograma'] = prazo3
    dfgeral.loc[3, 'Competência Base'] = hab_100
    dfgeral.loc[3, 'Competência Menor Desempenho'] = hab_100_geral
    dfgeral.loc[3, 'HardSkill'] = hardskill100
    dfgeral.loc[3, 'Desempenho'] = hardskill100
    dfgeral.loc[3, 'Certificações'] = certifica100
    dfgeral.loc[3, 'Cronograma'] = prazo4
    dfgeral.fillna(0)
    AgGrid(dfgeral)
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name=name)
        workbook = writer.book
        worksheet = writer.sheets[name]
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

    





