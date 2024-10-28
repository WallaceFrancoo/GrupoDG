#Bem-vindo

import pandas as pd
import calendar
import BancoDeDados
from tkinter import filedialog, Button, Label, Tk, messagebox, simpledialog
from datetime import datetime
import re
import os
from unidecode import unidecode
import PyPDF2

EmpresaDoAplicativo = ''
arquivoEmpresa = []
arquivoCompliance = []
arquivoChempack = []
arquivoCertificadora = []
def linhaArquivoEmpresa(data, debito, credito, valor, historico):
    limparcaracter = unidecode(historico)
    historicoAjustado = re.sub(r'[^\w\s/;.-]', '', limparcaracter)
    linhaCompleta = f"{data};{debito};{credito};{valor};;{historicoAjustado};1;;;"
    print(linhaCompleta)
def verificarEmpresaDestino(empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento):

    limparcaracter = unidecode(HistoricoCompleto)
    historicoAjustado = re.sub(r'[^\w\s/;.-]', '', limparcaracter)

    if empresa == '1ª DG':
        if debito == ContaBanco:
            credito = '803'
        elif credito == ContaBanco:
            debito = '803'
    elif empresa == '2ª CERTIFICADORA':
        if debito == ContaBanco:
            credito = '743'
        elif credito == ContaBanco:
            debito = '743'
    elif empresa == '3ª CHEMPACK':
        if debito == ContaBanco:
            credito = '578'
        elif credito == ContaBanco:
            debito = '578'
    linhaArquivoEmpresa(data,debito,credito,valorLancamento,HistoricoCompleto)
    linhaArquivoEntreEmpresas(empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento)

def linhaArquivoEntreEmpresas(empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento):
    if empresa == '1ª DG':
        if debito == '803':
            debito = BancoDeDados.ConsultarCompliance(lancamento)
            credito = '714'
        elif credito == '803':
            debito = '714'
            credito = BancoDeDados.ConsultarCompliance(lancamento)
        linhaArquivoCompliance(data, debito, credito, valorLancamento, HistoricoCompleto)
    elif empresa == '2ª CERTIFICADORA':
        if debito == ContaBanco:
            credito = '743'
        elif credito == ContaBanco:
            debito = '743'
    elif empresa == '3ª CHEMPACK':
        if debito == ContaBanco:
            credito = '578'
        elif credito == ContaBanco:
            debito = '578'

def linhaArquivoCompliance(data, debito, credito, valor, historico):
    linhaCompleta = f"{data};{debito};{credito};{valor};;{historico};1;;;"
    print(linhaCompleta)
def linhaArquivoChempack(empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco):
    print("Montar a linha para pagamentos para a Chempack")
def linhaArquivoCertificadora(empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco):
    print("Montar a linha para pagamentos para a Certificadora")
def fazerProcv(concatenacao, planilhaID, lancamento,ContaBanco, data, valor):

    # Lendo a planilha de identificação
    identificacao = pd.read_excel(planilhaID)

    # Verificando se o DataFrame não está vazio
    if not identificacao.empty:
        # Removendo espaços extras nos nomes das colunas
        identificacao.columns = identificacao.columns.str.strip()
        # Verificando se as colunas existem
        if 'DATA DO PAGTO/CRÉDITO' in identificacao.columns and 'VALOR PAGO/RECEBIDO' in identificacao.columns:
            # Convertendo a coluna de data para o formato desejado
            identificacao['DATA DO PAGTO/CRÉDITO'] = pd.to_datetime(identificacao['DATA DO PAGTO/CRÉDITO'],
                                                                    errors='coerce')
            identificacao['DATA DO PAGTO/CRÉDITO'] = identificacao['DATA DO PAGTO/CRÉDITO'].dt.strftime('%d/%m/%Y')

            # Concatenando as colunas 'DATA DO PAGTO/CRÉDITO' e 'VALOR PAGO/RECEBIDO'
            identificacao['concat_coluna'] = (
                    identificacao['DATA DO PAGTO/CRÉDITO'].astype(str).str.strip() +
                    identificacao['VALOR PAGO/RECEBIDO'].astype(str).str.strip()
            )
            # Verificando se o valor concatenado está presente na planilha
            if concatenacao in identificacao['concat_coluna'].values:
                # Recuperando a linha onde o valor foi encontrado
                complemento = ''
                linha_encontrada = identificacao[identificacao['concat_coluna'] == concatenacao]

                # Pegando o valor da coluna 'EMPRESA' para essa linha
                empresa_destino = linha_encontrada['EMPRESA'].values[0]  # Pega o valor da primeira correspondência
                tipo_documento = linha_encontrada['TIPO DOC'].values[0]
                referencia = linha_encontrada['REFERÊNCIA'].values[0]
                complemento = ''
                if tipo_documento in ['FOLHA', 'DANFE', 'NFS', 'IMPOSTO', 'NFE', 'NF-e', 'NFs', 'NF']:
                        complemento = linha_encontrada['Nº  DOCUMENTO'].fillna('').values[0]

                if empresa_destino in ['4ª DG - SERVIÇOS','5ª SASC','6ª JASC']:
                    print('Ele irá fazer o procv!')
                    debito, credito, valorLancamento, inicioHistorico  = verificarValor(valor, lancamento, data, ContaBanco)
                    HistoricoCompleto = f"{inicioHistorico}{referencia} - {complemento}"
                    linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)

                else:
                    print('Pagamento para outras empresas')
                    debito, credito, valorLancamento, inicioHistorico  = verificarValor(valor, lancamento, data, ContaBanco)
                    HistoricoCompleto = f"{inicioHistorico}{referencia} - {complemento}"
                    verificarEmpresaDestino(empresa_destino, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento)

                print(f'Achou! A empresa correspondente é: {empresa_destino}\n---------------------------------------------------------')
                return empresa_destino  # Retorna a informação encontrada
            else:
                print(f'Não Achou. Valores concatenados: {identificacao["concat_coluna"].values}')
                contraPartida = BancoDeDados.ConsultarDG(lancamento)
                print(contraPartida)
                print(f"Lançamento que não achou, mas tem no banco!\n--------------------------------------------------")



                return False
        else:
            print(
                f"O Valor é: {concatenacao}\nAs colunas 'DATA DO PAGTO/CRÉDITO' ou 'VALOR PAGO/RECEBIDO' não foram encontradas.")
            return False
    else:
        print("O DataFrame está vazio!")
        return False
def converter_valor(valor):
    try:
        negativo = '-' in valor
        valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
        valor_convertido = float(valor)
        return -valor_convertido if negativo else valor_convertido
    except:
        return 0
def is_valid_date_for_month(data, user_input):
    # Dicionário para converter meses em nomes para números
    months_map = {month: str(index).zfill(2) for index, month in enumerate(calendar.month_name) if month}
    months_map.update({month: str(index).zfill(2) for index, month in enumerate(calendar.month_abbr) if month})

    # Verificar se o user_input é número ou nome do mês
    if user_input.isdigit():
        month = str(int(user_input)).zfill(2)  # Converte para 06 se for 6 ou 06
    else:
        month = months_map.get(user_input.capitalize(), None)  # Verifica se é "junho", "Junho", etc.
        if month is None:
            return False  # Mês inválido

    try:
            # Tenta converter a data
        parsed_date = pd.to_datetime(data, format='%d/%m/%Y', errors='raise')
        # Verifica se o mês da data coincide com o mês do user_input
        if parsed_date.strftime('%m') == month:
            return True
        else:
            return False
    except (ValueError, TypeError):
        return False

def verificarValor(valor, lancamento, data, ContaBanco):
    debito = ''
    credito = ''
    if valor < 0 :
        debito = BancoDeDados.ConsultarDG(lancamento)
        credito = ContaBanco
        inicioHistorico = "PAGAMENTO REF. "
    elif valor > 0:
        debito = ContaBanco
        credito = BancoDeDados.ConsultarDG(lancamento)
        inicioHistorico = "RECEBIMENTO REF. "
        if credito == "187":
            if int(data[:2]) > 0 and int(data[:2]) < 8:
                credito = '187'
            elif int(data[:2]) > 8 and int(data[:2]) < 13:
                credito = '189'
            elif int(data[:2]) > 13 and int(data[:2]) < 21:
                credito = '25'
            else:
                credito = '187'
    valorLancamento = valor
    return debito, credito, valorLancamento, inicioHistorico
def processar_extrato(caminho_extrato,caminho_identificacao ,mes):
    ContaBanco = '537'

    try:
        extrato = pd.read_excel(caminho_extrato, header=8, engine='xlrd')  # Para arquivos XLS
    except ValueError:
        extrato = pd.read_excel(caminho_extrato, header=8, engine='openpyxl')  # Para arquivos XLSX

    try:
        identificacao = pd.read_excel(caminho_identificacao, header=0, engine='xlrd')  # Para arquivos XLS
    except ValueError:
        identificacao = pd.read_excel(caminho_identificacao, header=0, engine='openpyxl')  # Para arquivos XLSX


    extrato.columns = extrato.columns.str.strip()
    identificacao.columns = identificacao.columns.str.strip()

    colunas_extrato = ['Data', 'Lançamento', 'Crédito (R$)', 'Débito (R$)']
    colunas_identificacao = ['BANCO DO DEBITO/CRÉDITO']

    # Converter os valores para numéricos
    extrato['Crédito (R$)'] = extrato['Crédito (R$)'].astype(str).apply(converter_valor)
    extrato['Débito (R$)'] = extrato['Débito (R$)'].astype(str).apply(converter_valor)

    # Filtrar pelo mês
    extrato_filtrado = extrato[extrato['Data'].apply(is_valid_date_for_month, args=(mes,))]

    historicos = []
    for _, linha in extrato_filtrado.iterrows():
        data = linha['Data']
        lancamento = linha['Lançamento']
        credito = linha['Crédito (R$)']
        debito = linha['Débito (R$)']

        # Tratando NaN para a lógica do valor
        credito_valor = 0 if pd.isna(credito) else credito
        debito_valor = 0 if pd.isna(debito) else debito

        # Lógica para definir o valor
        if credito_valor > 0:
            valor = credito_valor
        elif debito_valor > 0:
            valor = -debito_valor
        else:
            continue  # Pular a linha se não houver valor

        # Substituindo NaN por string em branco para exibição
        credito_str = '' if pd.isna(credito) else str(credito_valor)
        debito_str = '' if pd.isna(debito) else str(debito_valor)

        # Concatenando as informações corretamente
        concatenacao = f"{data}{credito_str}{debito_str}".strip()  # Adiciona espaços e remove excessos

        # Imprimindo a concatenação
        print(concatenacao)

        # Passando a concatenação para a função
        fazerProcv(concatenacao, caminho_identificacao,lancamento,ContaBanco, data, valor)





caminhoE = 'Z:\Diversos\Projetos T.I\GrupoDG\JANEIRO DG SERVICOS.XLS'
caminhoI = 'Z:/Diversos/Projetos T.I/GrupoDG/fluxoDG.xlsx'
mes = '01'
processar_extrato(caminhoE,caminhoI,mes)