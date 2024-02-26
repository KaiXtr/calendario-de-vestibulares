from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import webbrowser
import pandas as pd
import sys
import re

class CalendarioCustomizado(QCalendarWidget):
    def __init__(self, *args):
        QCalendarWidget.__init__(self, *args)
        self.corInicio = QColor.fromRgb(10,200,10)
        self.corInicio.setAlpha(40)
        self.corFim = QColor.fromRgb(200,10,10)
        self.corFim.setAlpha(40)
        self.corPagamento = QColor.fromRgb(200,200,10)
        self.corPagamento.setAlpha(40)
        self.corProva = QColor.fromRgb(10,100,200)
        self.corProva.setAlpha(40)

        self.selectionChanged.connect(self.updateCells)
        self.diasInicio = []
        self.diasFim = []
        self.diasPagamento = []
        self.diasProva = []
    
    def adicionarDiaMarcado(self,dia,mes,tipo):
        dd = self.selectedDate()
        dd.setDate(dd.year(),mes,dia)
        if tipo == 0: self.diasInicio.append(dd)
        if tipo == 1: self.diasFim.append(dd)
        if tipo == 2: self.diasPagamento.append(dd)
        if tipo == 3: self.diasProva.append(dd)
  
    def paintCell(self, painter, rect, date):
        QCalendarWidget.paintCell(self, painter, rect, date)
        
        if date in self.diasInicio:
            painter.fillRect(rect, self.corInicio)
        if date in self.diasFim:
            painter.fillRect(rect, self.corFim)
        if date in self.diasPagamento:
            painter.fillRect(rect, self.corPagamento)
        if date in self.diasProva:
            painter.fillRect(rect, self.corProva)

class Informacao(QWidget):
    def __init__(self,v):
        super().__init__()
        self.vestibular = v
        self.meses_nomes = [
            "janeiro","fevereiro","março",
            "abril","maio","junho",
            "julho","agosto","setembro",
            "outubro","novembro","dezembro"
            ]

        self.setGeometry(100,100,500,600)

        self.label = QLabel(self)
        txt = v["nome"] + "\n"
        if "status" in v.keys() and v["status"] != "":
            txt += "Status: " + str(v["status"]) + "\n"
        if "inicio" in v.keys() and v["inicio"] != "":
            txt += "Início das inscrições: " + str(v["inicio"][0]) + " de " + self.meses_nomes[v["inicio"][1] - 1] + "\n"
        if "fim" in v.keys() and v["fim"] != "":
            txt += "Fim das inscrições: " + str(v["fim"][0]) + " de " + self.meses_nomes[v["fim"][1] - 1] + "\n"
        if "pagamento" in v.keys() and v["pagamento"] != "":
            txt += "Fim do pagamento: " + str(v["pagamento"][0]) + " de " + self.meses_nomes[v["pagamento"][1] - 1] + "\n"
        if "prova" in v.keys() and v["prova"] != "":
            txt += "Data da prova: " + str(v["prova"][0]) + " de " + self.meses_nomes[v["prova"][1] - 1] + "\n\n"
        if "obras" in v.keys() and str(v["obras"]) != "":
            txt += "Obras literárias: \n" + str(v["obras"]) + "\n\n"
        if "obs" in v.keys() and v["obs"] != "":
            txt += "Observações: \n" + str(v["obs"])
        self.label.setText(txt)
        self.label.setGeometry(100,10,400,400)

        self.edital = QPushButton(self)
        self.edital.setText("Abrir Edital")
        self.edital.clicked.connect(self.abrirLink)
        self.edital.setGeometry(100,450,100,50)

        self.fechar = QPushButton(self)
        self.fechar.setText("Fechar")
        self.fechar.clicked.connect(self.close)
        self.fechar.setGeometry(300,450,100,50)

        self.show()
    
    def abrirLink(self):
        webbrowser.open(self.vestibular["edital"])

class JanelaCalendario(QMainWindow):
    def __init__(self):
        super().__init__()
        self.vestibulares = []
        self.meses_nomes = [
            "janeiro","fevereiro","março",
            "abril","maio","junho",
            "julho","agosto","setembro",
            "outubro","novembro","dezembro"
            ]

        # Boas Vindas
        msg = QMessageBox()
        msg.setText("Para iniciar, indique a planilha contendo as informações.")
        msg.setWindowTitle("Iniciar")
        msg.exec_()

        # Abrir arquivo
        arquivo = ""
        while (arquivo.endswith(".xlsx") == False):
            arquivo = QFileDialog.getOpenFileName(self,"Abrir planilha...", ".", "*.xlsx")[0]
            if arquivo == "": exit()
            elif arquivo.endswith(".xlsx") == False:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Formato de arquivo inválido.")
                msg.setWindowTitle("Erro")
                msg.exec_()
        self.tabela = pd.read_excel(arquivo)
                
        # Janela
        self.setWindowTitle("Calendário de vestibulares")
        self.setGeometry(100, 100, 700, 450)

        # Calendário
        self.calendario = CalendarioCustomizado(self)
        self.calendario.setGeometry(50, 50, 400, 250)
        self.calendario.clicked.connect(self.atualizar)
        self.calendario.setMinimumDate(QDate(2020, 1, 1))
        self.calendario.setMaximumDate(QDate(2028, 1, 1))
        self.calendario.showToday()

        # Lista de concursos
        self.lista = QListWidget(self)
        self.lista.setGeometry(450, 50, 200, 250)
        self.lista.itemClicked.connect(self.redirecionar)

        # Filtrar data
        self.label = QLabel(self)
        self.label.setText("Selecionar data de:")
        self.label.setGeometry(450, 300, 200, 50)

        rbtxts = ("Início","Fim","Pagamento","Prova")
        self.radio = []
        for i in range(4):
            self.radio.append(QRadioButton(self))
            self.radio[i].setText(rbtxts[i])
            self.radio[i].setGeometry(450 + ((i%2) * 120), 350 + (int(i/2) * 30), 200, 20)
            if i == 0: self.radio[i].setChecked(True)
            if i == 0: self.radio[i].setToolTip("Data de início das inscrições")
            if i == 1: self.radio[i].setToolTip("Data de término das inscrições")
            if i == 2: self.radio[i].setToolTip("Data limite para pagamento das inscrições")
            if i == 3: self.radio[i].setToolTip("Data da prova")

        # Texto de eventos
        self.tblabel = QLabel(self)
        self.tblabel.setGeometry(50, 300, 100, 50)
        self.tblabel.setText("Eventos:")

        self.tbl = QListWidget(self)
        self.tbl.setGeometry(50, 350, 380, 80)
        self.tbl.itemClicked.connect(self.abrirInfo)

        # Estilos
        self.setStyleSheet(""
            + "QLabel#texto-eventos{"
            + "   background-color: #F2F2F2;"
            + "}"                 
            + "")

        # Exibir
        self.show()
        self.carregar()
    
    def carregar(self):
        col = 0
        for i in self.tabela.items(): #colunas
            lin = 0
            for j in i[1][2::]: #registros da coluna
                
                if col == 1: #obter nome do vestibular
                    self.vestibulares.append({})
                    self.vestibulares[lin]["id"] = lin
                    self.vestibulares[lin]["nome"] = j

                if col == 4: #obter status do vestibular
                    self.vestibulares[lin]["status"] = j
                    li = QListWidgetItem(self.vestibulares[lin]["nome"], self.lista)
                    
                    if self.vestibulares[lin]["status"] == "Inscrições Encerradas":
                        bb = QtGui.QBrush(QtGui.QColor(200,150,150))
                        li.setBackground(bb)

                if col == 6: #obter forma do vestibular
                    self.vestibulares[lin]["forma"] = j

                #obter datas do vestibular
                if col in (2,3,5,7) and j != "":
                    #múltiplas datas em uma célula
                    if re.compile("[A-Z]* - [0-9]*/[0-9]*/[0-9]*[A-z.\n\s-]*").match(str(j)):
                        jj = re.findall("[0-9]*/[0-9]*/[0-9]*",str(j))[0]
                        jj = jj.replace("/","-")
                        dti = pd.to_datetime(jj, format="%d-%m-%Y")
                    
                    #uma única data em uma única célula
                    else:
                        try: dti = pd.to_datetime(j, format="%d-%m-%Y")
                        except: continue

                    try:
                        dd = (int(dti.day),int(dti.month))

                        #classificar qual tipo de data é
                        if col == 2:
                            self.vestibulares[lin]["inicio"] = dd
                            self.calendario.adicionarDiaMarcado(int(dti.day),int(dti.month),0)
                        if col == 3:
                            self.vestibulares[lin]["fim"] = dd
                            self.calendario.adicionarDiaMarcado(int(dti.day),int(dti.month),1)
                        if col == 5:
                            self.vestibulares[lin]["pagamento"] = dd
                            self.calendario.adicionarDiaMarcado(int(dti.day),int(dti.month),2)
                        if col == 7:
                            self.vestibulares[lin]["prova"] = dd
                            self.calendario.adicionarDiaMarcado(int(dti.day),int(dti.month),3)

                    except:
                        print(f"Erro: não foi possível obter dados do registro \"{j}\"")
                
                #Obter link do edital
                if col == 8:
                    self.vestibulares[lin]["edital"] = j
                
                #Obter obras literárias
                if col == 9:
                    self.vestibulares[lin]["obras"] = j

                #Obter observações
                if col == 10:
                    self.vestibulares[lin]["obs"] = j

                lin += 1
            col += 1
    
    def atualizar(self):
        datual = self.calendario.selectedDate()
        chave = ("inicio","fim","pagamento","prova")
        tipo = ("início da inscrição","fim da inscrição","fim do pagamento","data da prova")
        self.tbl.clear()
        
        for i in self.vestibulares:
            #comparar data selecionada e datas da planilha
            for j in range(len(tipo)):
                if ((chave[j] in i.keys()) and (i[chave[j]][0] == int(datual.day()))
                and (i[chave[j]][1] == int(datual.month()))):
                    dados = i["nome"] + " - " + tipo[j] + ": "
                    dados += str(i[chave[j]][0]) + " de " + self.meses_nomes[i[chave[j]][1] - 1]
                    li = QListWidgetItem(dados, self.tbl)
                    li.id = i["id"]

    def redirecionar(self,item):
        datual = self.calendario.selectedDate()

        if self.radio[0].isChecked(): chave = "inicio"
        if self.radio[1].isChecked(): chave = "fim"
        if self.radio[2].isChecked(): chave = "pagamento"
        if self.radio[3].isChecked(): chave = "prova"

        for i in self.vestibulares:
            if (chave in i.keys()) and (i["nome"] == item.text()):
                datual.setDate(datual.year(),i[chave][1],i[chave][0])
                self.calendario.setSelectedDate(datual)
                self.atualizar()

    def abrirInfo(self,item):
        for i in self.vestibulares:
            if i["id"] == item.id:
                self.popup = Informacao(i)

if __name__ == "__main__":
    App = QApplication(sys.argv)
    window = JanelaCalendario()
    sys.exit(App.exec())