import sys
from PyQt5 import QtWidgets, uic, QtCore, QtGui, QtPrintSupport
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDesktopServices
from PyQt5.QtCore import Qt
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter, QPrintPreviewDialog
from PyQt5.QtWidgets import (QApplication, QTreeWidget, QTreeWidgetItem, QDialog, QPushButton, QFileDialog, QMessageBox, QToolBar)
from tools import * #mensajebox, preguntar, v_titulonav, textolabelEditar, textolabelNuevo, nombreiconoNuevo, nombreiconoEditar
import xlwt
from datetime import date, datetime
import json
import math

v_titulo = "Informe de Ventas"

with open('sesion.json', 'r') as f: #Carga el JSON con el usuario
    cadena_json = json.load(f)
    idusuario = cadena_json['id']
    usuario = cadena_json['usuario']


querysql_select = """
select
    fv.idFactura,
    fv.idTimbrado,
    t.Timbrado,
    fv.NroFactura,
    fv.FechaFactura,
    fv.idClientes,
    c.razonsocial,
    c.ruc_ci,
    fv.idUsuarios, 
    us.usuario,
    (SELECT SUM(PVenta * cantidad) FROM ventas_det where idFactura = fv.idFactura) as total
    from ventas_encab fv
        inner join clientes c on fv.idClientes = c.idClientes
        inner join timbrado t on fv.idTimbrado = t.idTimbrado
        inner join usuarios us on fv.IdUsuarios = us.IdUsuarios
        inner join ventas_det vd on fv.idFactura = vd.idFactura
	group by fv.idFactura
    order by fv.idFactura
"""
querysql_select_2 = """
select
    fv.idFactura,
    fv.idTimbrado,
    t.Timbrado,
    fv.NroFactura,
    fv.FechaFactura,
    fv.idClientes,
    c.razonsocial,
    c.ruc_ci,
    fv.idUsuarios, 
    us.usuario,
    (SELECT SUM(PVenta * cantidad) FROM ventas_det where idFactura = fv.idFactura) as total
    from ventas_encab fv
        inner join clientes c on fv.idClientes = c.idClientes
        inner join timbrado t on fv.idTimbrado = t.idTimbrado
        inner join usuarios us on fv.IdUsuarios = us.IdUsuarios
        inner join ventas_det vd on fv.idFactura = vd.idFactura
    where fv.FechaFactura between '{} 00:00:00'  AND '{} 23:59:59' 
	group by fv.idFactura
    order by fv.idFactura
"""

querysql_selectmax = "select ifnull(max(idFactura), 0) + 1 from ventas_encab"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba



etiquetascol = ["Código Factura", "Código Timbrado", "Timbrado", "Nro Factura", "Fecha de Emisión", "idCliente", "Razón Social", "RUC/CI", "idUsuario", "Usuario", "Total de Venta"]



class iniciar:
    def __init__(self, conexion, listaapegar = None):
        self.mylistaapegar = listaapegar
        self.conex = conexion
        cursor = self.conex.cursor() #Valida el permiso
        querysql_permiso = "select valor from usuariopermisos where idusuarios like "+idusuario+" and idpermisos like 7"
        cursor.execute(querysql_permiso)
        self.permiso = cursor.fetchone()
        
        self.vnav = uic.loadUi("win_filtro_ventas.ui")

        self.vnav.label_2.setText("Ingresos Totales: ")
        self.vnav.label.setText("IVA 10%: ")
        self.vnav.label_3.setText("IVA 5%: ")
        self.vnav.label_4.setText("Fecha Desde: ")
        self.vnav.label_5.setText("Fecha Hasta: ")

        self.vnav.lineEdit_3.setReadOnly(True)
        self.vnav.lineEdit_2.setReadOnly(True)
        self.vnav.lineEdit.setReadOnly(True)
        vfecha = QtCore.QDate.currentDate()
        self.vnav.dateEdit.setDate(vfecha)
        self.vnav.dateEdit_2.setDate(vfecha)
        
        self.cargardatos()

        self.vnav.tableWidget.setHorizontalHeaderLabels(etiquetascol) #Establece texto en las etiquetas
        self.vnav.tableWidget.resizeColumnsToContents() #Reajusta columnas a su contenido

        if self.mylistaapegar is None:
            self.vnav.setWindowTitle(v_titulo)
            self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo))


        self.vnav.showMaximized()
        self.vnav.show()
        self.vnav.actionCerrar.triggered.connect(self.click_actioncerrar)
        self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
        self.vnav.actionExcel.triggered.connect(self.click_actionExcel)
        self.vnav.pushButton.clicked.connect(self.clicked_pushButton)
        #self.vnav.tableWidget.activated.connect(self.activated_tableWidget)
        #self.vnav.actionInforme.triggered.connect(self.click_actionInforme)
        

    def clicked_pushButton(self): #Busca por fecha las ventas
        fecha1 = self.vnav.dateEdit.dateTime().toString("yyyy-MM-dd")
        fecha2 = self.vnav.dateEdit_2.dateTime().toString("yyyy-MM-dd")
        cursor = self.conex.cursor()
        cursor.execute(querysql_select_2.format(fecha1, fecha2))

        fila = 0
        self.vnav.tableWidget.clearContents()
        self.vnav.tableWidget.setRowCount(0)
        for campoid, campodescrip, campo3, campo4, campo5, \
            campo8, campo9, campo10, campo11, campo12, campototal in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(str(campodescrip))
            cel2 = QtWidgets.QTableWidgetItem(campo3)
            cel3 = QtWidgets.QTableWidgetItem(str(campo4))
            cel4 = QtWidgets.QTableWidgetItem(str(campo5))
            
            cel7 = QtWidgets.QTableWidgetItem(str(campo8))
            cel8 = QtWidgets.QTableWidgetItem(campo9)
            cel9 = QtWidgets.QTableWidgetItem(campo10)
            cel10 = QtWidgets.QTableWidgetItem(str(campo11))
            cel11 = QtWidgets.QTableWidgetItem(campo12)
            cel12 = QtWidgets.QTableWidgetItem(str(campototal))

            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel11.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel12.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3)
            self.vnav.tableWidget.setItem(fila, 4, cel4) 
            self.vnav.tableWidget.setItem(fila, 5, cel7)
            self.vnav.tableWidget.setItem(fila, 6, cel8)
            self.vnav.tableWidget.setItem(fila, 7, cel9)
            self.vnav.tableWidget.setItem(fila, 8, cel10)
            self.vnav.tableWidget.setItem(fila, 9, cel11)
            self.vnav.tableWidget.setItem(fila, 10, cel12)
            #
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))


        subtotal = 0
        iva = 0
        iva5 = 0
        nb_row = self.vnav.tableWidget.rowCount()
        for row in range (nb_row):
            colum5 = self.vnav.tableWidget.item(row, 10).text()
            subtotal = subtotal + float(colum5)

        total = subtotal
        iva5 = 0
        iva10 = total / 11
        self.vnav.lineEdit_2.setText(str(math.trunc(total)))
        self.vnav.lineEdit.setText(str(math.trunc(iva10)))
        self.vnav.lineEdit_3.setText(str(iva5))


    def click_actioncancelar(self):
        self.vedit.close()

    

    def click_actionImprimir(self):
        fecha1 = self.vnav.dateEdit.dateTime().toString("dd-MM-yyyy")
        fecha2 = self.vnav.dateEdit_2.dateTime().toString("dd-MM-yyyy")
        v0 = self.vnav.lineEdit_2.text()
        v1 = self.vnav.lineEdit.text()
        v2 = self.vnav.lineEdit_3.text()
        if self.cantfilas > 0:
            now = datetime.now()
            vvfecha = now.strftime("%d-%m-%Y, %H-%M-%S")
            filename = "Informe ventas {}.pdf".format(str(vvfecha)) #Nombre del archivo
            model = self.vnav.tableWidget.model() #Apuntador al tableWidget

            printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.PrinterResolution)
            printer.setOutputFormat(QtPrintSupport.QPrinter.PdfFormat)
            printer.setPaperSize(QtPrintSupport.QPrinter.A4)
            printer.setOrientation(QtPrintSupport.QPrinter.Landscape)
            printer.setOutputFileName(filename)

            doc = QtGui.QTextDocument()

            html = """<html>
            <head>
            <style>
            table, th, td {
            border: 1px solid black;
            border-collapse: collapse;
            }
            
            </style>
            </head>"""
            html += "<h1>{}</h1><br>".format(v_titulo)
            html += "<table><thead>" #La tabla para reporte
            html += "<tr>"
            for c in range(model.columnCount()):
                html += "<th>{}</th>".format(model.headerData(c, QtCore.Qt.Horizontal))

            html += "</tr></thead>"
            html += "<tbody>"
            for r in range(model.rowCount()):
                html += "<tr>"
                for c in range(model.columnCount()):
                    html += "<td>{}</td>".format(model.index(r, c).data() or "")
                html += "</tr>"
            html += "</tbody></table>"
            html += """
                    <br><br>
                    <b>Fecha Desde: </b>{} <br />
                    <b>Fecha Hasta: </b>{} <br />
                    <b>Total de Ingresos: </b>{} Gs.<br />
                    <b>IVA 10%: </b>{} Gs.<br />
                    <b>IVA 5%: </b>{} Gs.
                    """.format(fecha1, fecha2, v0, v1, v2)
            doc.setHtml(html)
            doc.setPageSize(QtCore.QSizeF(printer.pageRect().size()))
            doc.print_(printer)
            QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(filename)) #Abre el PDF automatico
        else:
            self.vnav.statusbar.showMessage("No hay registros para exportar")

    def click_actionExcel(self): #Exportar a Excel la tabla
        if self.cantfilas > 0:
            now = datetime.now()
            vvfecha = now.strftime("%d-%m-%Y, %H-%M-%S")
            filename = QtWidgets.QFileDialog.getSaveFileName(self.vnav, 'Exportar a Excel', '{} {}'.format(v_titulo, str(vvfecha)), ".xls(*.xls)")
            if len(filename[0]) > 0:
                libro = xlwt.Workbook()
                hoja1 = libro.add_sheet("hoja1", cell_overwrite_ok=True)
                hoja1.write(0, 0, v_titulo)
                model = self.vnav.tableWidget.model()
                for currentColumn in range(self.vnav.tableWidget.columnCount()):
                    encabezadolabel = model.headerData(currentColumn, QtCore.Qt.Horizontal)
                    hoja1.write(1, currentColumn, encabezadolabel)
                    for currentRow in range(self.vnav.tableWidget.rowCount()):
                        textocelda = str(self.vnav.tableWidget.item(currentRow, currentColumn).text())
                        hoja1.write(currentRow+2, currentColumn, textocelda)
                libro.save(filename[0])
                QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(filename[0])) #Abre el archivo automaticamente
        else:
            self.vnav.statusbar.showMessage("No hay registros para exportar")


    def click_actioncerrar(self):
        self.vnav.close()

    #Cargar datos y mostrar en las tablas
    def cargardatos(self):
        cursor = self.conex.cursor()
        cursor.execute(querysql_select)

        fila = 0
        self.vnav.tableWidget.clearContents()
        self.vnav.tableWidget.setRowCount(0)
        for campoid, campodescrip, campo3, campo4, campo5, \
            campo8, campo9, campo10, campo11, campo12, campototal in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(str(campodescrip))
            cel2 = QtWidgets.QTableWidgetItem(campo3)
            cel3 = QtWidgets.QTableWidgetItem(str(campo4))
            cel4 = QtWidgets.QTableWidgetItem(str(campo5))
            
            cel7 = QtWidgets.QTableWidgetItem(str(campo8))
            cel8 = QtWidgets.QTableWidgetItem(campo9)
            cel9 = QtWidgets.QTableWidgetItem(campo10)
            cel10 = QtWidgets.QTableWidgetItem(str(campo11))
            cel11 = QtWidgets.QTableWidgetItem(campo12)
            cel12 = QtWidgets.QTableWidgetItem(str(campototal))

            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel11.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel12.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3)
            self.vnav.tableWidget.setItem(fila, 4, cel4) 
            self.vnav.tableWidget.setItem(fila, 5, cel7)
            self.vnav.tableWidget.setItem(fila, 6, cel8)
            self.vnav.tableWidget.setItem(fila, 7, cel9)
            self.vnav.tableWidget.setItem(fila, 8, cel10)
            self.vnav.tableWidget.setItem(fila, 9, cel11)
            self.vnav.tableWidget.setItem(fila, 10, cel12)
            #
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))