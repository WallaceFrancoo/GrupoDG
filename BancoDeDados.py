import sqlite3
import datetime
from tkinter import messagebox

conn = sqlite3.connect('Z:/Diversos/Projetos T.I/GrupoDG/DePara/ConceptaDGServico.db')
cursor = conn.cursor()

cursor.execute('''
CREATE TABLE IF NOT EXISTS deParaDGServicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    DE TEXT NOT NULL,
    PARA TEXT NOT NULL
)
''')
def ConsultarDG(valor):
    cursor.execute('SELECT * FROM deParaDGServicos')
    tabelaHistorico = cursor.fetchall()

    if tabelaHistorico:
        cursor.execute('SELECT DE, DGSERVICOS FROM deParaDGServicos WHERE ? LIKE "%" || DE || "%" COLLATE NOCASE', (valor,))
        tabela = cursor.fetchall()
        if tabela:
            return tabela[0][1]
        else:
            return f"6"
    else:
        return f"6"
def ConsultarCompliance(valor):
    cursor.execute('SELECT * FROM deParaDGServicos')
    tabelaHistorico = cursor.fetchall()

    if tabelaHistorico:
        cursor.execute('SELECT DE, DGCOMPLIANCE FROM deParaDGServicos WHERE ? LIKE "%" || DE || "%" COLLATE NOCASE',
                       (valor,))
        tabela = cursor.fetchall()
        if tabela:
            return tabela[0][1]
        else:
            return f"6"
    else:
        return f"6"
def ConsultarCertificadora(valor):
    cursor.execute('SELECT * FROM deParaDGServicos')
    tabelaHistorico = cursor.fetchall()

    if tabelaHistorico:
        cursor.execute('SELECT DE, CERTIFICADORA FROM deParaDGServicos WHERE ? LIKE "%" || DE || "%" COLLATE NOCASE',
                       (valor,))
        tabela = cursor.fetchall()
        if tabela:
            return tabela[0][1]
        else:
            return f"6"
    else:
        return f"6"
def ConsultarChempack(valor):
    cursor.execute('SELECT * FROM deParaDGServicos')
    tabelaHistorico = cursor.fetchall()

    if tabelaHistorico:
        cursor.execute('SELECT DE, CHEMPACK FROM deParaDGServicos WHERE ? LIKE "%" || DE || "%" COLLATE NOCASE',
                       (valor,))
        tabela = cursor.fetchall()
        if tabela:
            return tabela[0][1]
        else:
            return f"6"
    else:
        return f"6"
def adicionar_DadosChempack(DE, PARA):
    cursor.execute('SELECT * FROM deParaDGServicos WHERE DE = ? ',(DE,))
    resultado = cursor.fetchone()

    if resultado:
        retorno = resultado[4]
        resposta = messagebox.askquestion("Atualizar", f"Essa associação DE-PARA já existe na DGServicos com a conta {retorno}. \nDeseja atualizar os dados?", icon='warning')

        if resposta == 'yes':
            cursor.execute('UPDATE deParaDGServicos SET CHEMPACK = ? WHERE DE = ?', (PARA, DE))
            conn.commit()
            messagebox.showinfo("Sucesso","Dados atualizados com sucesso")
        else:
            messagebox.showinfo("Aviso", "Os Dados nao foram atualizados")
    else:
        cursor.execute('INSERT INTO deParaDGServicos(DE,CHEMPACK) VALUES (?, ?)',(DE, PARA))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Cadastro do historico {DE}, foi cadastrado para a conta {PARA}!")

def adicionar_DadosCertificadora(DE, PARA):
    cursor.execute('SELECT * FROM deParaDGServicos WHERE DE = ? ',(DE,))
    resultado = cursor.fetchone()

    if resultado:
        retorno = resultado[3]
        resposta = messagebox.askquestion("Atualizar", f"Essa associação DE-PARA já existe na Certificadora com a conta {retorno}. \nDeseja atualizar os dados?", icon='warning')

        if resposta == 'yes':
            cursor.execute('UPDATE deParaDGServicos SET CERTIFICADORA = ? WHERE DE = ?', (PARA, DE))
            conn.commit()
            messagebox.showinfo("Sucesso","Dados atualizados com sucesso")
        else:
            messagebox.showinfo("Aviso", "Os Dados nao foram atualizados")
    else:
        cursor.execute('INSERT INTO deParaDGServicos(DE,CERTIFICADORA) VALUES (?, ?)',(DE, PARA))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Cadastro do historico {DE}, foi cadastrado para a conta {PARA}!")

def adicionar_DadosDGompliance(DE, PARA):
    cursor.execute('SELECT * FROM deParaDGServicos WHERE DE = ? ',(DE,))
    resultado = cursor.fetchone()

    if resultado:
        retorno = resultado[2]
        resposta = messagebox.askquestion("Atualizar", f"Essa associação DE-PARA já existe na DGCOMPLIANCE com a conta {retorno}. \nDeseja atualizar os dados?", icon='warning')

        if resposta == 'yes':
            cursor.execute('UPDATE deParaDGServicos SET DGCOMPLIANCE = ? WHERE DE = ?', (PARA, DE))
            conn.commit()
            messagebox.showinfo("Sucesso","Dados atualizados com sucesso")
        else:
            messagebox.showinfo("Aviso", "Os Dados nao foram atualizados")
    else:
        cursor.execute('INSERT INTO deParaDGServicos(DE, DGCOMPLIANCE) VALUES (?, ?)',(DE, PARA))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Cadastro do historico {DE}, foi cadastrado para a conta {PARA}!")

def adicionar_DadosDGServicos(DE, PARA):
    cursor.execute('SELECT * FROM deParaDGServicos WHERE DE = ? ',(DE,))
    resultado = cursor.fetchone()

    if resultado:
        retorno = resultado[1]
        resposta = messagebox.askquestion("Atualizar", f"Essa associação DE-PARA já existe na DGServicos com a conta {retorno}. \nDeseja atualizar os dados?", icon='warning')

        if resposta == 'yes':
            cursor.execute('UPDATE deParaDGServicos SET DGSERVICOS = ? WHERE DE = ?', (PARA, DE))
            conn.commit()
            messagebox.showinfo("Sucesso","Dados atualizados com sucesso")
        else:
            messagebox.showinfo("Aviso", "Os Dados nao foram atualizados")
    else:
        cursor.execute('INSERT INTO deParaDGServicos(DE,DGSERVICOS) VALUES (?, ?)',(DE, PARA))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Cadastro do historico {DE}, foi cadastrado para a conta {PARA}!")

def adicionar_DadosGeral(DE, PARA):
    cursor.execute('SELECT * FROM deParaDGServicos WHERE DE = ? ',(DE,))
    resultado = cursor.fetchone()

    if resultado:
        retorno = resultado[1]
        resposta = messagebox.askquestion("Atualizar", f"Essa associação DE-PARA já existe na DGServicos com a conta {retorno}. \nDeseja atualizar os dados?", icon='warning')

        if resposta == 'yes':
            cursor.execute('UPDATE deParaDGServicos SET DGSERVICOS = ? SET DGCOMPLIANCE = ? SET CERTIFICADORA = ? SET CHEMPACK = ? WHERE DE = ?', (PARA, PARA, PARA, PARA, DE))
            conn.commit()
            messagebox.showinfo("Sucesso","Dados atualizados com sucesso")
        else:
            messagebox.showinfo("Aviso", "Os Dados nao foram atualizados")
    else:
        cursor.execute('INSERT INTO deParaDGServicos(DE,DGSERVICOS,DGCOMPLIANCE,CERTIFICADORA,CHEMPACK) VALUES (?, ?, ?, ?, ?)',(DE, PARA, PARA, PARA, PARA))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Cadastro do historico {DE}, foi cadastrado para a conta {PARA}!")
