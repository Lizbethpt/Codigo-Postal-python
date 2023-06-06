import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTreeWidgetItem, QTreeWidget
from PyQt5 import QtCore, QtGui, QtWidgets
import win32com.client as win32
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from openpyxl import Workbook
# Ejemplo de uso
from docx import Document

from openpyxl import Workbook


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.treeWidget = QTreeWidget(self.centralwidget)
        self.treeWidget.setGeometry(QtCore.QRect(10, 160, 281, 241))
        self.treeWidget.setObjectName("treeWidget")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(20, 10, 93, 28))
        self.pushButton.setObjectName("pushButton")

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(150, 10, 141, 28))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(30, 50, 93, 28))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(160, 50, 161, 28))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QtCore.QRect(610, 100, 93, 28))
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setGeometry(QtCore.QRect(610, 140, 93, 28))
        self.pushButton_7.setObjectName("pushButton_7")
        self.treeView_2 = QtWidgets.QTreeView(self.centralwidget)
        self.treeView_2.setGeometry(QtCore.QRect(340, 360, 256, 192))
        self.treeView_2.setObjectName("treeView_2")
        self.pushButton_8 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_8.setGeometry(QtCore.QRect(650, 370, 93, 28))
        self.pushButton_8.setObjectName("pushButton_8")
        self.pushButton_9 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_9.setGeometry(QtCore.QRect(650, 410, 93, 28))
        self.pushButton_9.setObjectName("pushButton_9")
        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView.setGeometry(QtCore.QRect(340, 90, 256, 192))
        self.graphicsView.setObjectName("graphicsView")

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Codigo postales"))
        self.pushButton.setText(_translate("MainWindow", "Abrir"))

        self.pushButton_3.setText(_translate("MainWindow", "guardar con open xml"))
        self.pushButton_4.setText(_translate("MainWindow", "Guardar pdf"))
        self.pushButton_5.setText(_translate("MainWindow", "guardar  en word"))
        self.pushButton_6.setText(_translate("MainWindow", "Grafcar "))
        self.pushButton_7.setText(_translate("MainWindow", "Limpiar"))
        self.pushButton_8.setText(_translate("MainWindow", "Abrir arbol"))
        self.pushButton_9.setText(_translate("MainWindow", "Cerrar"))


class MiVentana(QMainWindow):
    d1 = []

    def __init__(self):
        super().__init__()

        # Configurar la interfaz gráfica generada por pyuic5
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.cargar_archivo)
        self.ui.pushButton_4.clicked.connect(self.pdfscript)
        self.ui.pushButton_5.clicked.connect(self.guardar_en_documento)
        self.ui.pushButton_3.clicked.connect(self.guardar_en_excel)

    def guardar_en_documento(self):
        document = Document()

        for row in self.d1:
            table = document.add_table(rows=1, cols=len(row))
            table.autofit = False

            for i, cell in enumerate(row):
                table.cell(0, i).text = str(cell)

        document.save('datos.docx')

    def guardar_en_excel(self):
        workbook = Workbook()
        sheet = workbook.active

        for row_data in self.d1:
            sheet.append(row_data)

        workbook.save('datos.xlsx')

    def pdfscript(self):

        # Datos de ejemplo
        data = self.d1
        pdf_filename = 'tabla_datos.pdf'

        custom_size = (8 * 180, 80 * 20)  # Convertir pulgadas a puntos (1 pulgada = 72 puntos)
        pdf = SimpleDocTemplate(pdf_filename, pagesize=custom_size)
        # Crear la tabla
        table = Table(data)

        # Estilo de la tabla
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        # Aplicar el estilo a la tabla
        table.setStyle(style)

        # Crear la lista de elementos a agregar al PDF
        elements = [table]

        # Generar el PDF
        pdf.build(elements)

        print(f"El archivo '{pdf_filename}' ha sido creado.")

    def cargar_archivo(self):
        ruta_archivo, _ = QFileDialog.getOpenFileName(None, "Seleccionar archivo", "", "Archivos de texto (*.txt)")

        with open(ruta_archivo, 'r', encoding='cp1252') as archivo:
            columnas = archivo.readline().strip()
            columna = columnas.split('|')
            self.d1.append(columna)
            # Configurar las columnas del QTreeWidget
            self.ui.treeWidget.setColumnCount(len(columna))
            self.ui.treeWidget.setHeaderLabels(columna)

            for linea in archivo:
                datos = linea.strip().split('|')

                item = QTreeWidgetItem(self.ui.treeWidget, datos)

                self.d1.append(datos)

                self.ui.treeWidget.addTopLevelItem(item)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ventana = MiVentana()
    ventana.show()
    sys.exit(app.exec_())
