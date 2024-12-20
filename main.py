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
arquivoDGServicos = []
relatorioErro = []
mes = []
BancoDGServicos = 'BRADESCO DG SERVIÇOS'
BancoDGCompliance = ['BRADESCO DG','BRADESCO DG COMPLIANCE']
BancoChempack = 'BRADESCO CHEMPACK'
BancoCertificadora = 'BRADESCO CERTIFICADORA'


def linhaArquivoEmpresa(data, debito, credito, valor, historico):
    limparcaracter = unidecode(historico)
    historicoAjustado = re.sub(r'[^\w\s/;.-]', '', limparcaracter)
    valorNovo = str(valor).replace("-", "").replace(".", ",")
    linhaCompleta = f"{data};{debito};{credito};{valorNovo};;{historicoAjustado};1;;;"

    arquivoEmpresa.append({
        'data': data,
        'lançamento': historicoAjustado,
        'valor': valorNovo,
        'Debito': debito,
        'Credito': credito
    })
    print(f"{debito} - {credito}")

    if str(debito).strip() == "6" or str(credito).strip() == "6":
        print('vai ir pro relatorio')
        relatorioErro.append({
            'data': data,
            'lançamento': historicoAjustado,
            'valor': valorNovo,
            'Debito': debito,
            'Credito': credito
        })

    print(linhaCompleta)
def verificarEmpresaDestino(opcao, banco, empresaPagante, empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento):

    limparcaracter = unidecode(HistoricoCompleto)
    historicoAjustado = re.sub(r'[^\w\s/;.-]', '', limparcaracter)
    print(opcao)
    if opcao == '1ª DG': #842
        if empresa == '3ª CHEMPACK':
            if debito == ContaBanco:
                credito = '803'
            elif credito == ContaBanco:
                debito = '803'
        elif empresa == '2ª CERTIFICADORA':
            if debito == ContaBanco:
                credito = '743'
            elif credito == ContaBanco:
                debito = '743'
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == ContaBanco:
                credito = '714'
            elif credito == ContaBanco:
                debito = '714'
    elif opcao == '2ª CERTIFICADORA': #841
        if empresa == '3ª CHEMPACK':
            if debito == ContaBanco:
                credito = '578'
            elif credito == ContaBanco:
                debito = '578'
        elif empresa == '1ª DG':
            if debito == ContaBanco:
                credito = '726'
            elif credito == ContaBanco:
                debito = '726'
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == ContaBanco:
                credito = '581'
            elif credito == ContaBanco:
                debito = '581'
    elif opcao == '3ª CHEMPACK': # 840
        if empresa == '2ª CERTIFICADORA':
            if debito == ContaBanco:
                credito = '725'
            elif credito == ContaBanco:
                debito = '725'
        elif empresa == '1ª DG':
            if debito == ContaBanco:
                credito = '712'
            elif credito == ContaBanco:
                debito = '712'
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == ContaBanco:
                credito = '726'
            elif credito == ContaBanco:
                debito = '726'
    elif opcao == '4ª DG - SERVIÇOS': #843
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
    linhaArquivoEntreEmpresas(opcao, empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento)
def linhaArquivoEntreEmpresas(opcao, empresa, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento):
    if opcao == '4ª DG - SERVIÇOS':
        if empresa == '1ª DG':
            if debito == '803':
                debito = BancoDeDados.ConsultarCompliance(lancamento)
                credito = '714'
            elif credito == '803':
                debito = '714'
                credito = BancoDeDados.ConsultarCompliance(lancamento)
            linhaArquivoCompliance(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '2ª CERTIFICADORA':
            if debito == '743':
                debito = BancoDeDados.ConsultarCertificadora(lancamento)
                credito = '581'
            elif credito == '743':
                debito = '581'
                credito = BancoDeDados.ConsultarCertificadora(lancamento)
            linhaArquivoCertificadora(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '3ª CHEMPACK':
            if debito == '578':
                debito = BancoDeDados.ConsultarChempack(lancamento)
                credito = '726'
            elif credito == '578':
                debito = '726'
                credito = BancoDeDados.ConsultarChempack(lancamento)
            linhaArquivoChempack(data, debito, credito, valorLancamento, HistoricoCompleto)
    elif opcao == '3ª CHEMPACK':
        if empresa == '1ª DG':
            if debito == '712':
                debito = BancoDeDados.ConsultarCompliance(lancamento)
                credito = '803'
            elif credito == '712':
                debito = '803'
                credito = BancoDeDados.ConsultarCompliance(lancamento)
            linhaArquivoChempack(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '2ª CERTIFICADORA':
            if debito == '725':
                debito = BancoDeDados.ConsultarCertificadora(lancamento)
                credito = '578'
            elif credito == '725':
                debito = '578'
                credito = BancoDeDados.ConsultarCertificadora(lancamento)
            linhaArquivoCertificadora(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == '726':
                debito = BancoDeDados.ConsultarDG(lancamento)
                credito = '578'
            elif credito == '726':
                debito = '578'
                credito = BancoDeDados.ConsultarDG(lancamento)
            linhaArquivoDGServicos(data, debito, credito, valorLancamento, HistoricoCompleto)
    elif opcao == '2ª CERTIFICADORA':
        if empresa == '1ª DG':
            if debito == '726':
                debito = BancoDeDados.ConsultarCompliance(lancamento)
                credito = '743'
            elif credito == '726':
                debito = '743'
                credito = BancoDeDados.ConsultarCompliance(lancamento)
        elif empresa == '3ª CHEMPACK':
            if debito == '578':
                debito = BancoDeDados.ConsultarChempack(lancamento)
                credito = '725'
            elif credito == '578':
                debito = '725'
                credito = BancoDeDados.ConsultarChempack(lancamento)
            linhaArquivoChempack(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == '581':
                debito = BancoDeDados.ConsultarDG(lancamento)
                credito = '743'
            elif credito == '581':
                debito = '743'
                credito = BancoDeDados.ConsultarDG(lancamento)
            linhaArquivoDGServicos(data, debito, credito, valorLancamento, HistoricoCompleto)
    elif opcao == '1ª DG':
        if empresa == '3ª CHEMPACK':
            if debito == '803':
                debito = BancoDeDados.ConsultarChempack(lancamento)
                credito = '712'
            elif credito == '803':
                debito = '712'
                credito = BancoDeDados.ConsultarChempack(lancamento)
            linhaArquivoChempack(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '4ª DG - SERVIÇOS':
            if debito == '714':
                debito = BancoDeDados.ConsultarDG(lancamento)
                credito = '803'
            elif credito == '714':
                debito = '803'
                credito = BancoDeDados.ConsultarDG(lancamento)
            linhaArquivoDGServicos(data, debito, credito, valorLancamento, HistoricoCompleto)
        elif empresa == '2ª CERTIFICADORA':
            if debito == '743':
                debito = BancoDeDados.ConsultarCertificadora(lancamento)
                credito = '726'
            elif credito == '743':
                debito = '726'
                credito = BancoDeDados.ConsultarCertificadora(lancamento)
            linhaArquivoCertificadora(data, debito, credito, valorLancamento, HistoricoCompleto)
def linhaArquivoCompliance(data, debito, credito, valor, historico):

    valorNovo = str(valor).replace("-", "").replace(".", ",")
    arquivoCompliance.append({
        'data': data,
        'lançamento': historico,
        'valor': valorNovo,
        'Debito': debito,
        'Credito': credito
    })
    if debito == "6" or credito == "6":
        relatorioErro.append({
            'data': data,
            'lançamento': historico,
            'valor': valorNovo,
            'Debito': debito,
            'Credito': credito
        })
    print('Executou a linha Compliance')
def linhaArquivoChempack(data, debito, credito, valor, historico):
    valorNovo = str(valor).replace("-", "").replace(".", ",")
    arquivoChempack.append({
        'data': data,
        'lançamento': historico,
        'valor': valorNovo,
        'Debito': debito,
        'Credito': credito
    })
    if debito == "6" or credito == "6":
        relatorioErro.append({
            'data': data,
            'lançamento': historico,
            'valor': valorNovo,
            'Debito': debito,
            'Credito': credito
        })
    print('Executou a linha Chempack')
def linhaArquivoCertificadora(data, debito, credito, valor, historico):
    valorNovo = str(valor).replace("-", "").replace(".", ",")
    arquivoCertificadora.append({
        'data': data,
        'lançamento': historico,
        'valor': valorNovo,
        'Debito': debito,
        'Credito': credito
    })
    if debito == "6" or credito == "6":
        relatorioErro.append({
            'data': data,
            'lançamento': historico,
            'valor': valorNovo,
            'Debito': debito,
            'Credito': credito
        })
    print('Executou a linha Certificadora')
def linhaArquivoDGServicos(data, debito, credito, valor, historico):

    valorNovo = str(valor).replace("-", "").replace(".", ",")
    arquivoDGServicos.append({
        'data': data,
        'lançamento': historico,
        'valor': valorNovo,
        'Debito': debito,
        'Credito': credito
    })
    if debito == "6" or credito == "6":
        relatorioErro.append({
            'data': data,
            'lançamento': historico,
            'valor': valorNovo,
            'Debito': debito,
            'Credito': credito
        })

    print('Executou a linha DGServicos')
def fazerProcv(banco, opcao, concatenacao, planilhaID, lancamento,ContaBanco, data, valor):
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
                bancoPagamento = linha_encontrada['BANCO DO DEBITO/CRÉDITO'].values[0]
                empresa_destino = linha_encontrada['EMPRESA'].values[0]  # Pega o valor da primeira correspondência
                tipo_documento = linha_encontrada['TIPO DOC'].values[0]
                referencia = linha_encontrada['REFERÊNCIA'].values[0]
                complemento = ''
                if bancoPagamento in banco:
                    if tipo_documento in ['FOLHA', 'DANFE', 'NFS', 'IMPOSTO', 'NFE', 'NF-e', 'NFs', 'NF']:
                            complemento = linha_encontrada['Nº  DOCUMENTO'].fillna('').values[0]

                    if empresa_destino in [opcao ,'5ª SASC','6ª JASC']:
                        print(f'Ele irá fazer o procv!{opcao}')
                        debito, credito, valorLancamento, inicioHistorico  = verificarValor(opcao, valor, lancamento, data, ContaBanco)
                        HistoricoCompleto = f"{inicioHistorico}{referencia} - {complemento}"
                        linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)
                    elif tipo_documento == 'MÚTUO':
                        print('Emprestimo entre empresas!')
                        debito, credito, valorLancamento, inicioHistorico = verificarValor(opcao, valor, lancamento,
                                                                                           data,
                                                                                           ContaBanco)
                        HistoricoCompleto = f"{inicioHistorico}{lancamento}"
                        linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)
                    else:
                        print(f'Pagamento para outras empresas {opcao}')
                        debito, credito, valorLancamento, inicioHistorico  = verificarValor(opcao, valor, lancamento, data, ContaBanco)
                        HistoricoCompleto = f"{inicioHistorico}{referencia} - {complemento}"
                        verificarEmpresaDestino(opcao, banco, bancoPagamento, empresa_destino, data, debito, credito, valorLancamento, HistoricoCompleto, ContaBanco, lancamento)
                    print(f'Achou! A empresa correspondente é: {empresa_destino}\n---------------------------------------------------------')
                    return empresa_destino  # Retorna a informação encontrada
                else:
                    print("revisar aqui")
                    debito, credito, valorLancamento, inicioHistorico = verificarValor(opcao, valor, lancamento, data,
                                                                                       ContaBanco)
                    HistoricoCompleto = f"{inicioHistorico}{lancamento}"
                    linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)

            else:
                print(f'Não Achou. Valores concatenados: {identificacao["concat_coluna"].values}')
                debito, credito, valorLancamento, inicioHistorico = verificarValor(opcao, valor, lancamento, data, ContaBanco)
                HistoricoCompleto = f"{inicioHistorico}{lancamento}"
                linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)
                print(f"Lançamento que não achou, mas tem no banco!\n--------------------------------------------------")
                return False
        else:
            print(
                f"O Valor é: {concatenacao}\nAs colunas 'DATA DO PAGTO/CRÉDITO' ou 'VALOR PAGO/RECEBIDO' não foram encontradas.")
            debito, credito, valorLancamento, inicioHistorico = verificarValor(opcao, valor, lancamento, data, ContaBanco)
            HistoricoCompleto = f"{inicioHistorico}{lancamento}"
            linhaArquivoEmpresa(data, debito, credito, valor, HistoricoCompleto)
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
def verificarValor(opcao, valor, lancamento, data, ContaBanco):
    debito = ''
    credito = ''
    if opcao == '4ª DG - SERVIÇOS':
        if valor < 0 :
            debito = BancoDeDados.ConsultarDG(lancamento)
            credito = ContaBanco
            inicioHistorico = "PAGAMENTO REF. "
        elif valor > 0:
            debito = ContaBanco
            credito = BancoDeDados.ConsultarDG(lancamento)
            inicioHistorico = "RECEBIMENTO REF. "
    elif opcao == '1ª DG':
        if valor < 0 :
            debito = BancoDeDados.ConsultarCompliance(lancamento)
            credito = ContaBanco
            inicioHistorico = "PAGAMENTO REF. "
        elif valor > 0:
            debito = ContaBanco
            credito = BancoDeDados.ConsultarCompliance(lancamento)
            inicioHistorico = "RECEBIMENTO REF. "
    elif opcao == '3ª CHEMPACK':
        if valor < 0 :
            debito = BancoDeDados.ConsultarChempack(lancamento)
            credito = ContaBanco
            inicioHistorico = "PAGAMENTO REF. "
        elif valor > 0:
            debito = ContaBanco
            credito = BancoDeDados.ConsultarChempack(lancamento)
            inicioHistorico = "RECEBIMENTO REF. "
    elif opcao == '2ª CERTIFICADORA':
        if valor < 0 :
            debito = BancoDeDados.ConsultarCertificadora(lancamento)
            credito = ContaBanco
            inicioHistorico = "PAGAMENTO REF. "
        elif valor > 0:
            debito = ContaBanco
            credito = BancoDeDados.ConsultarCertificadora(lancamento)
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
def processar_extrato(banco, opcao, caminho_extrato,caminho_identificacao ,mes):
    ContaBanco = '537'
    print(f"Verificando arquivo: {caminho_extrato}")
    extensao = caminho_extrato.split('.')[-1].lower()  # Convertendo para minúsculas
    print(f"Extensão do arquivo: {extensao}")

    try:
        # Verificar a extensão do arquivo e escolher o engine adequado
        if extensao == 'xls':
            extrato = pd.read_excel(caminho_extrato, header=8, engine='xlrd')  # Para arquivos .xls
        elif extensao == 'xlsx':
            extrato = pd.read_excel(caminho_extrato, header=8, engine='openpyxl')  # Para arquivos .xlsx
        else:
            raise ValueError(f"Formato de arquivo não suportado: {caminho_extrato}")
    except Exception as e:
        print(f"Erro ao ler o arquivo de extrato: {e}")
        return  # Retorna se houver erro

    try:
        if caminho_identificacao.endswith('.xls'):
            identificacao = pd.read_excel(caminho_identificacao, header=0, engine='xlrd')  # Para arquivos .xls
        elif caminho_identificacao.endswith('.xlsx'):
            identificacao = pd.read_excel(caminho_identificacao, header=0, engine='openpyxl')  # Para arquivos .xlsx
        else:
            raise ValueError(f"Formato de arquivo não suportado: {caminho_identificacao}")
    except Exception as e:
        print(f"Erro ao ler o arquivo de identificação: {e}")
        # Opcionalmente, logue o erro ou forneça um feedback ao usuário

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
        print(opcao)

        # Passando a concatenação para a função
        fazerProcv(banco, opcao, concatenacao, caminho_identificacao,lancamento,ContaBanco, data, valor)

    return arquivoEmpresa, arquivoCompliance, arquivoChempack, arquivoCertificadora, arquivoDGServicos, relatorioErro
def buscarArquivo(opcao):
    global arquivoEmpresa, arquivoCompliance, arquivoChempack, arquivoCertificadora, arquivoDGServicos, relatorioErro, mes

    mes = simpledialog.askstring("Entrada", "Digite o mês (Numerico!):")
    arquivo = filedialog.askopenfilename(
        title="Selecione o Arquivo Bradesco",
        filetypes=(("Planilhas", "*.*"), ("Todos os arquivos", "*.*"))
    )
    arquivoIdentificacao = filedialog.askopenfilename(
        title="Selecione a planilha de identificacao!",
        filetypes=(("Identificação","*.*"), ("Todos os arquivos", "*.*"))
    )

    BancoDGServicos = 'BRADESCO DG SERVIÇOS'
    BancoDGCompliance = ['BRADESCO DG', 'BRADESCO DG COMPLIANCE']
    BancoChempack = 'BRADESCO CHEMPACK'
    BancoCertificadora = 'BRADESCO CERTIFICADORA'
    banco = ''

    if opcao == '840':
        opcao = '3ª CHEMPACK'
        banco = BancoChempack
    elif opcao == '841':
        opcao = '2ª CERTIFICADORA'
        banco = BancoCertificadora
    elif opcao == '842':
        opcao = '1ª DG'
        banco = BancoDGCompliance
    elif opcao == '843':
        opcao = '4ª DG - SERVIÇOS'
        banco = BancoDGServicos

    if arquivo:
        arquivoEmpresa, arquivoCompliance, arquivoChempack, arquivoCertificadora, arquivoDGServicos, relatorioErro = processar_extrato(banco, opcao, arquivo, arquivoIdentificacao,mes)
        messagebox.showinfo('Sucesso','Os documentos estão prontos para serem gerados!')
        return
    else:
        return None
def gerarArquivo(empresa):
    global arquivoEmpresa, arquivoCompliance, arquivoChempack, arquivoCertificadora, arquivoDGServicos, relatorioErro
    messagebox.showinfo('Aviso', 'Irá gerar os arquivos para importação em cada empresa')

    caminho_arquivo_txt = filedialog.asksaveasfilename(
        title="Salvar arquivo como",
        defaultextension=".txt",
        initialfile=f"ARQUIVO PRINCIPAL DA EMPRESA - {empresa} - {mes}",
        filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
    )
    if caminho_arquivo_txt:
        try:
            # Arquivo principal
            with open(caminho_arquivo_txt, "w") as arquivoPrincipal:
                for dado in arquivoEmpresa:
                    linha_formatada = (f"{dado['data']};"
                                       f"{dado['Debito']};"
                                       f"{dado['Credito']};"
                                       f"{dado['valor']};;"
                                       f"{dado['lançamento']};1;;; ")
                    arquivoPrincipal.write(linha_formatada.upper() + '\n')
                messagebox.showinfo('Sucesso', f"Arquivo geral salvo com sucesso!")

            # Arquivo Compliance
            caminho_Compliance_txt = filedialog.asksaveasfilename(
                title="Salvar relatório da complince TXT",
                defaultextension=".txt",
                initialfile=f"842 - PAGAMENTOS PARA A COMPLIANCE DA EMPRESA - {empresa} - {mes}",
                filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
            )
            if caminho_Compliance_txt:
                with open(caminho_Compliance_txt, "w") as arquivo2:
                    for dado in arquivoCompliance:
                        linha_formatada = (f"{dado['data']};"
                                           f"{dado['Debito']};"
                                           f"{dado['Credito']};"
                                           f"{dado['valor']};;"
                                           f"{dado['lançamento']};1;;; ")
                        arquivo2.write(linha_formatada.upper() + '\n')
                messagebox.showinfo('Sucesso', f"Arquivo Compliance salvo com sucesso!")

            # Arquivo Chempack
            caminho_Chempack_txt = filedialog.asksaveasfilename(
                title="Salvar relatório da Chempack TXT",
                defaultextension=".txt",
                initialfile=f"840 - PAGAMENTOS PARA A CHEMPACK DA EMPRESA - {empresa} - {mes}",
                filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
            )
            if caminho_Chempack_txt:
                with open(caminho_Chempack_txt, "w") as arquivo3:
                    for dado in arquivoChempack:
                        linha_formatada = (f"{dado['data']};"
                                           f"{dado['Debito']};"
                                           f"{dado['Credito']};"
                                           f"{dado['valor']};;"
                                           f"{dado['lançamento']};1;;; ")
                        arquivo3.write(linha_formatada.upper() + '\n')
                messagebox.showinfo('Sucesso', f"Arquivo Chempack salvo com sucesso!")

            # Arquivo Certificadora
            caminho_Certificadora_txt = filedialog.asksaveasfilename(
                title="Salvar relatório da Certificadora TXT",
                defaultextension=".txt",
                initialfile=f"841 - PAGAMENTOS PARA A CERTIFICADORA DA EMPRESA - {empresa} - {mes}",
                filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
            )
            if caminho_Certificadora_txt:
                with open(caminho_Certificadora_txt, "w") as arquivo4:
                    for dado in arquivoCertificadora:
                        linha_formatada = (f"{dado['data']};"
                                           f"{dado['Debito']};"
                                           f"{dado['Credito']};"
                                           f"{dado['valor']};;"
                                           f"{dado['lançamento']};1;;; ")
                        arquivo4.write(linha_formatada.upper() + '\n')
                messagebox.showinfo('Sucesso', f"Arquivo Certificadora salvo com sucesso!")

                # Arquivo DGServicos
            caminho_DGServicos_txt = filedialog.asksaveasfilename(
                title="Salvar relatório da Certificadora TXT",
                defaultextension=".txt",
                initialfile=f"843 - PAGAMENTOS PARA A DG SERVICOS DA EMPRESA - {empresa} - {mes}",
                filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
                )
            if caminho_DGServicos_txt:
                with open(caminho_DGServicos_txt, "w") as arquivo5:
                    for dado in arquivoCertificadora:
                        linha_formatada = (f"{dado['data']};"
                                            f"{dado['Debito']};"
                                            f"{dado['Credito']};"
                                            f"{dado['valor']};;"
                                            f"{dado['lançamento']};1;;; ")
                        arquivo5.write(linha_formatada.upper() + '\n')
                messagebox.showinfo('Sucesso', f"Arquivo DG Servicos salvo com sucesso!")
            # Relatório de erros
            caminho_relatorio_txt = filedialog.asksaveasfilename(
                title="Salvar relatório de erros como",
                defaultextension=".txt",
                initialfile=f"Contas em Acerto - {mes}",
                filetypes=(("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
            )
            if caminho_relatorio_txt:
                with open(caminho_relatorio_txt, "w") as relatorio_arquivo:
                    for relatorio in relatorioErro:
                        linha_formatada = (f"{relatorio['data']};{relatorio['lançamento']};{relatorio['valor']};"
                                           f"{relatorio['Debito']};{relatorio['Credito']}\n")
                        relatorio_arquivo.write(linha_formatada.upper())
                messagebox.showinfo('Sucesso', f"Relatório de erros salvo com sucesso!")
                os.startfile(caminho_relatorio_txt)

        except Exception as e:
            print(f"Erro ao salvar arquivo TXT: {e}")

        # Zera as variáveis globais para evitar duplicação
        arquivoEmpresa = []
        arquivoCompliance = []
        arquivoChempack = []
        arquivoCertificadora = []
        arquivoDGServicos = []
        relatorioErro = []

    else:
        print('Local de salvamento não selecionado')
