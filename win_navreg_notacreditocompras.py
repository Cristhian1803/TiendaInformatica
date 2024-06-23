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

v_titulo = "Nota de Crédito de Compras"

with open('sesion.json', 'r') as f: #Carga el JSON con el usuario
    cadena_json = json.load(f)
    idusuario = cadena_json['id']
    usuario = cadena_json['usuario']

querysql_update = """
 update notacreditocompra set 
    idFactura = {},
    idTimbrado = {},
    NroFactura = {},
    FechaFactura = '{}',
    Condicion = '{}',
    Anulado = {},
    idProveedores = {},
    idTiposFacturas = {},
    idUsuarios = {}
      where idFactura = {}
"""
querysql_insert = """
 insert into notacreditocompra(
    idNotaCredCompra,
    NroFactura,
    Fecha,  
    idUsuarios)
     values(
     {}, 
     nullif(trim({}),''),  
     nullif(trim('{}'),''),
     {}                                
     )
"""

querysql_insert2 = """
 insert into detnotacredcompra(
    idNotaCredCompra,
    idProductos,
    PCompra,
    Cantidad)
     values(
     {}, 
     nullif(trim({}),''), 
     nullif(trim({}),''),  
     nullif(trim({}),'')                                  
     )
"""

querysql_delete = "delete from notacreditocompra where idnotacredcompra = {}"
querysql_delete2 = "delete from detnotacredcompra where idnotacredcompra = {}"

querysql_select = """
select
    nc.idNotaCredCompra,
    nc.NroFactura,
    fv.idProveedores,
    c.razonsocial,
    c.ruc_ci,
    nc.Fecha,
    nc.idUsuarios, 
    us.usuario from notacreditocompra nc
        inner join compras_encab fv on nc.NroFactura = fv.NroFactura
        inner join proveedores c on fv.idProveedores = c.idProveedores
        inner join usuarios us on fv.IdUsuarios = us.IdUsuarios
    order by nc.idNotaCredCompra
"""

querysql_select2 = """
select
    nc.idProductos,
    p.NombreProd,
    nc.PCompra,
    nc.Cantidad,
    nc.PCompra * nc.Cantidad as subtotal,
	nc.PorcIva 
    from compras_det nc
        inner join prodservicio p on nc.idProductos = p.idProductos
	where NroFactura = {}
    order by idProductos
"""
querysql_selectmax = "select ifnull(max(idNotaCredCompra), 0) + 1 from notacreditocompra"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba



etiquetascol = ["Codigo", "Nro de Factura", "idProveedor", "Razon Social", "RUC-CI", "Fecha", "idUsuario", "Usuario"]



class iniciar:
    def __init__(self, conexion, listaapegar = None):
        self.mylistaapegar = listaapegar
        self.conex = conexion
        cursor = self.conex.cursor() #Valida el permiso
        querysql_permiso = "select valor from usuariopermisos where idusuarios like "+idusuario+" and idpermisos like 7"
        cursor.execute(querysql_permiso)
        self.permiso = cursor.fetchone()
        
        self.vnav = uic.loadUi("win_navreg_notacreditocompras.ui")
        
        self.cargardatos()

        self.vnav.tableWidget.setHorizontalHeaderLabels(etiquetascol) #Establece texto en las etiquetas
        self.vnav.tableWidget.resizeColumnsToContents() #Reajusta columnas a su contenido

        if self.mylistaapegar is None:
            self.vnav.setStyleSheet("QMainWindow{background-color: " + colornormalventana +";}")
            self.vnav.setWindowTitle(v_titulo)
            self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo))
        else:
            self.vnav.setStyleSheet("QMainWindow{background-color: " + colorpegarventana +";}")
            self.vnav.setWindowTitle(v_titulo + " > MODO REFERENCIAL")
            self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo + "(REFERENCIAL)"))

        self.vnav.showMaximized()
        self.vnav.show()
        self.vnav.actionCerrar.triggered.connect(self.click_actioncerrar)
        self.vnav.actionNuevo.triggered.connect(self.click_actionNuevo)
        self.vnav.actionEditar.triggered.connect(self.click_actionEditar)
        self.vnav.actionEditar.setEnabled(False)
        self.vnav.actionEliminar.triggered.connect(self.click_actionEliminar)
        #self.vnav.actionEliminar.setEnabled(False)
        self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
        self.vnav.actionExcel.triggered.connect(self.click_actionExcel)
        #self.vnav.tableWidget.activated.connect(self.activated_tableWidget)
        

    def activated_tableWidget(self): #Se activa al darle Enter con una fila seleccionada
        if self.mylistaapegar is None:
            self.click_actionEditar()
        else:
            try:
                v0 = self.vnav.tableWidget.selectedIndexes()[0].data()
                v1 = self.vnav.tableWidget.selectedIndexes()[1].data()
                v2 = self.vnav.tableWidget.selectedIndexes()[2].data() #
                self.mylistaapegar[0].setText(v0)
                self.mylistaapegar[1].setText(v1)
                self.mylistaapegar[2].setText(v2) #     
                self.vnav.close()
            except:
                self.vnav.statusbar.showMessage("Seleccione un registro para pegar")

    def click_actionNuevo(self):
        self.editando = False
        self.cargarventanaedicion()
    
    def click_actionEditar(self):
        self.editando = True
        self.cargarventanaedicion()

    def cargarventanaedicion(self):
        #Carga la ventana
        try:

            if self.permiso == (1,):
                self.vedit = uic.loadUi("win_edicion_notacreditocompras.ui")
                self.vedit.label.setText(etiquetascol[0] + ": ")
                self.vedit.label_4.setText("Factura de Compra #: ")
                self.vedit.label_5.setText("Fecha: ")
                self.vedit.label_6.setText("Proveedor: ")
                self.vedit.label_9.setText("Usuario: ")
                self.vedit.label_10.setText("Producto: ")
                self.vedit.label_13.setText("Precio Unitario: ")
                self.vedit.label_14.setText("IVA (%): ")
                self.vedit.label_15.setText("Cantidad: ")
                self.vedit.label_18.setText("Precio: ")
                self.vedit.label_11.setText("Subtotal: ")
                self.vedit.label_16.setText("IVA10%: ")
                self.vedit.label_8.setText("IVA5%: ")
                self.vedit.label_17.setText("Total a Pagar: ")
                
                self.vedit.lineEdit.setReadOnly(True)
                #self.vedit.lineEdit_2.setReadOnly(True)
                #self.vedit.lineEdit_9.setReadOnly(True)
                self.vedit.lineEdit_3.setReadOnly(True)
                self.vedit.lineEdit_4.setReadOnly(True)
                self.vedit.lineEdit_5.setReadOnly(True)
                self.vedit.lineEdit_6.setReadOnly(True)
                self.vedit.lineEdit_12.setReadOnly(True)
                self.vedit.lineEdit_13.setReadOnly(True)
                self.vedit.lineEdit_14.setReadOnly(True)
                self.vedit.lineEdit_14.setVisible(False)
                self.vedit.lineEdit_16.setReadOnly(True)
                self.vedit.lineEdit_17.setReadOnly(True)
                self.vedit.lineEdit_18.setReadOnly(True)
                self.vedit.lineEdit_21.setReadOnly(True)
                self.vedit.lineEdit_22.setReadOnly(True)
                self.vedit.lineEdit_22.setVisible(False)
                self.vedit.lineEdit_23.setReadOnly(True)
                self.vedit.lineEdit_23.setVisible(False)
                self.vedit.lineEdit_24.setReadOnly(True)
                self.vedit.lineEdit_25.setReadOnly(True)
                self.vedit.dateTimeEdit.setReadOnly(True)
                self.vedit.lineEdit_24.setVisible(False)
                self.vedit.lineEdit_26.setVisible(False)
                self.vedit.lineEdit_10.setVisible(False)
                self.vedit.lineEdit_19.setReadOnly(True)
                self.vedit.lineEdit_11.setReadOnly(True)
                self.vedit.tableWidget2.clearContents()
                self.vedit.tableWidget2.setRowCount(0)
            else:
                self.mensaje = uic.loadUi("advertencia.ui")
                self.mensaje.lb_mensaje.setText("<html><head/><body><p><span style=' font-size:14pt; font-weight:600; color:#900000;'>Usted No tiene acceso a esta ventana</span></p></body></html>")
                resp = self.mensaje.exec()

            if self.editando:
                tituloventana = "Editar - {}".format(v_titulo)
                titulointerno = textolabelEditar.format(v_titulo)
                iconoVentana = QIcon(nombreiconoEditar)
                imagenventana = QPixmap(nombreiconoEditar)
                v0 = self.vnav.tableWidget.selectedIndexes()[0].data()
                v1 = self.vnav.tableWidget.selectedIndexes()[1].data() 
                v2 = self.vnav.tableWidget.selectedIndexes()[2].data()
                v3 = self.vnav.tableWidget.selectedIndexes()[3].data()
                v4 = self.vnav.tableWidget.selectedIndexes()[4].data()
                v5 = self.vnav.tableWidget.selectedIndexes()[5].data()#
                self.filaamodificar = self.vnav.tableWidget.currentRow() #Modifica en tiempo real la tabla
                vchk6 = self.vnav.tableWidget.item(self.filaamodificar, 6) #


                v7 = self.vnav.tableWidget.selectedIndexes()[7].data()
                v8 = self.vnav.tableWidget.selectedIndexes()[8].data()
                v9 = self.vnav.tableWidget.selectedIndexes()[9].data()
                v10 = self.vnav.tableWidget.selectedIndexes()[10].data()#

            else:
                tituloventana = "Nuevo - {}".format(v_titulo)
                titulointerno = textolabelNuevo.format(v_titulo)
                iconoVentana = QIcon(nombreiconoNuevo)
                imagenventana = QPixmap(nombreiconoNuevo)
                if querysql_selectmax is None:
                    v0 = ""
                else: #Sirve como semi autoincrementable
                    cursor = self.conex.cursor()
                    cursor.execute(querysql_selectmax)
                    res = cursor.fetchone()
                    v0 = str(res[0])
                    #v3 = str(res[0])
                    
               
                v1 = ""
                v2 = ""
                v3 = ""
                v4 = ""
                #v5 = "" 
                vchk6 = False 
                v7 = ""
                #v8 = ""
                v9 = ""
                v10 = ""
                v11 = ""
                v12 = ""
                v13 = "{}".format(idusuario)
                v14 = "{}".format(usuario)
                v15 = "" 
                v16 = ""
                v17 = ""
                v18 = "10"
                v19 = ""
                v20 = ""
                v21 = "" #
                self.filaamodificar = self.cantfilas

            #Carga las variables en LineEdit
            self.vedit.lineEdit.setText(v0)
            #self.vedit.lineEdit_2.setText(v1) 
            #self.vedit.lineEdit_9.setText(v2)
            vfecha = QtCore.QDateTime.currentDateTime()
            self.vedit.dateTimeEdit.setDateTime(vfecha) #fecha
            self.vedit.lineEdit_3.setText(v3)
            self.vedit.lineEdit_4.setText(v7)
            self.vedit.lineEdit_5.setText(v9)
            self.vedit.lineEdit_6.setText(v10)
            self.vedit.lineEdit_19.setText(v13)
            self.vedit.lineEdit_11.setText(v14)
            self.vedit.lineEdit_12.setText(v15)
            self.vedit.lineEdit_13.setText(v16)
            self.vedit.lineEdit_20.setText(v17)
            #self.vedit.lineEdit_14.setText(v18)
            self.vedit.lineEdit_15.setText(v19)
            self.vedit.lineEdit_21.setText(v20) #
            self.condicion_update = v0 #Captura el id para colocar el update en el mismo lugar en la BD

            self.vedit.setWindowTitle(tituloventana)
            self.vedit.setWindowIcon(iconoVentana)
            self.vedit.lb_titulo.setText(titulointerno)
            self.vedit.label_2.setPixmap(imagenventana)
            self.vedit.show() #Muestra la ventana

            self.vedit.actionAceptar.triggered.connect(self.click_actionaceptar)
            self.vedit.actionCancelar.triggered.connect(self.click_actioncancelar)

            self.vedit.pushButton.clicked.connect(self.clicked_pushButton)
            #self.vedit.pushButton_2.clicked.connect(self.clicked_pushButton_2)
            #self.vedit.pushButton_3.clicked.connect(self.clicked_pushButton_3)
            self.vedit.pushButton_4.clicked.connect(self.clicked_pushButton_4)
            self.vedit.pushButton_6.clicked.connect(self.calculateSubTotal)
            self.vedit.pushButton_5.clicked.connect(self.clicked_pushButton_5)
            self.vedit.pushButton_7.clicked.connect(self.calculateTotal)
            self.vedit.aggProd.clicked.connect(self.aggProd)

        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para Editar")


    def aggProd(self): #Importa los productos del Detalle de Factura
        try:
            #print("IMPORTAR")
            v0 = self.vedit.lineEdit_3.text()
            cursor = self.conex.cursor()       
            cursor.execute(querysql_select2.format(v0))

            fila = 0
            self.vedit.tableWidget2.clearContents()
            self.vedit.tableWidget2.setRowCount(0)
            for campoid, campodescrip, campo3, campo4, campo5, campo8 in cursor: #
                cel0 = QtWidgets.QTableWidgetItem(str(campoid))
                cel1 = QtWidgets.QTableWidgetItem(campodescrip)
                cel2 = QtWidgets.QTableWidgetItem(str(campo3))
                cel3 = QtWidgets.QTableWidgetItem(str(campo4))
                cel4 = QtWidgets.QTableWidgetItem(str(campo5))
                cel7 = QtWidgets.QTableWidgetItem(str(campo8))

                
                #Hace que las filas no sean editables
                cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                
                self.vedit.tableWidget2.insertRow(fila)
                self.vedit.tableWidget2.setItem(fila, 0, cel0)
                self.vedit.tableWidget2.setItem(fila, 1, cel1)
                self.vedit.tableWidget2.setItem(fila, 2, cel2)
                self.vedit.tableWidget2.setItem(fila, 3, cel3)
                self.vedit.tableWidget2.setItem(fila, 4, cel4) 
                self.vedit.tableWidget2.setItem(fila, 5, cel7)#
                fila += 1
        except:
            self.vedit.statusbar.showMessage("Ocurrió un Error. Recuerde agregar el Nro de Factura") 

    def clicked_pushButton_5(self): #Añade productos a la tabla de venta
        v15 = self.vedit.lineEdit_12.text()
        v16 = self.vedit.lineEdit_13.text()
        v17 = self.vedit.lineEdit_20.text()
        v18 = self.vedit.comboBox.currentText()
        v19 = self.vedit.lineEdit_15.text()
        v20 = self.vedit.lineEdit_21.text()

        contadorfila = self.vedit.tableWidget2.rowCount()
        vvfila = contadorfila
        if len(v15) > 0 and len(v16) > 0 and len(v17) > 0 and len(v18) > 0 and len(v19) > 0 and len(v20) > 0:
            celda00 = QtWidgets.QTableWidgetItem(v15)
            celda01 = QtWidgets.QTableWidgetItem(v16)
            celda02 = QtWidgets.QTableWidgetItem(v17)
            celda03 = QtWidgets.QTableWidgetItem(v18)
            celda04 = QtWidgets.QTableWidgetItem(v19)
            celda05 = QtWidgets.QTableWidgetItem(v20)
            self.vedit.tableWidget2.insertRow(vvfila)

            self.vedit.tableWidget2.setItem(vvfila, 0, celda00)
            self.vedit.tableWidget2.setItem(vvfila, 1, celda01)#
            self.vedit.tableWidget2.setItem(vvfila, 2, celda02)#
            self.vedit.tableWidget2.setItem(vvfila, 3, celda04)#
            self.vedit.tableWidget2.setItem(vvfila, 4, celda05)
            self.vedit.tableWidget2.setItem(vvfila, 5, celda03)
            vvfila+=1
            self.vedit.lineEdit_21.setText("")
            self.vedit.lineEdit_12.setText("")
            self.vedit.lineEdit_13.setText("")
            self.vedit.lineEdit_20.setText("")
            self.vedit.lineEdit_15.setText("")
            vvfila+=1
            self.vedit.lineEdit_21.setText("")
        else:
            self.vedit.statusbar.showMessage("Rellene todos los campos. Recuerde calcular el subtotal")


    def calculateSubTotal(self): #Calcula el subtotal
        rate = self.vedit.lineEdit_15.text()
        value = self.vedit.lineEdit_20.text()
        try:
            rate01 = float(rate)
            value01 = float(value)
            subtotal = value01 * rate01
            item_subtotal = self.vedit.lineEdit_21
            if item_subtotal is None:
                item_subtotal = 0
                self.vedit.lineEdit_15.setText(item_subtotal)

            item_subtotal.setText(str(subtotal))
        except:
            self.vedit.statusbar.showMessage("Introduzca valores numéricos en Precio y Cantidad")

    def calculateTotal(self): #Calcula el total
        subtotal = 0
        iva = 0
        iva5 = 0
        nb_row = self.vedit.tableWidget2.rowCount()
        for row in range (nb_row):
            colum4 = self.vedit.tableWidget2.item(row, 4).text()
            colum5 = self.vedit.tableWidget2.item(row, 5).text()
            subtotal = subtotal + float(colum4)
            if float(colum5) == 10:
                calciva10 = float(colum4) / 10
                iva = iva + calciva10
            else:
                calciva5 = float(colum4) / 5
                iva5 = iva5 + calciva5
        total = subtotal + iva + iva5

        item_subtotal = self.vedit.lineEdit_16
        item_iva = self.vedit.lineEdit_17
        item_total = self.vedit.lineEdit_18
        item_iva5 = self.vedit.lineEdit_25
        if item_subtotal is None:
            item_subtotal = 0
            self.vedit.lineEdit_16.setText(item_subtotal)
        if item_iva is None:
            item_iva = 0
            item_iva5 = 0
            self.vedit.lineEdit_17.setText(item_iva)
            self.vedit.lineEdit_25.setText(item_iva5)
        if item_total is None:
            item_total = 0
            self.vedit.lineEdit_18.setText(item_total)
        item_subtotal.setText(str(subtotal))
        item_iva.setText(str(iva))
        item_total.setText(str(total))
        item_iva5.setText(str(iva5))

    def clicked_pushButton(self):
        import win_navreg_compras
        self.listalineedit_pks_compra = [self.vedit.lineEdit_3, self.vedit.lineEdit_4, self.vedit.lineEdit_5, self.vedit.lineEdit_6]
        self.run_win_navreg_compras = win_navreg_compras.iniciar(self.conex, self.listalineedit_pks_compra)

    #def clicked_pushButton_3(self): #Busca tipoCobro
     #   import win_navreg_tipoCobro
      #  self.listalineedit_pks_tipco = [self.vedit.lineEdit_7, self.vedit.lineEdit_8]
       # self.run_win_navreg_tipoCobro = win_navreg_tipoCobro.iniciar(self.conex, self.listalineedit_pks_tipco)
        
    def clicked_pushButton_4(self): #Busca Items
        import win_navreg_items
        self.listalineedit_pks_items = [self.vedit.lineEdit_12, self.vedit.lineEdit_22, self.vedit.lineEdit_13, self.vedit.lineEdit_14, self.vedit.lineEdit_20]
        self.run_win_navreg_items = win_navreg_items.iniciar(self.conex, self.listalineedit_pks_items)

    def click_actionaceptar(self):
        v0 = self.vedit.lineEdit.text() #Lee los datos del LineEdit
        #v1 = self.vedit.lineEdit_2.text() #idTimbrado
        #v2 = self.vedit.lineEdit_9.text()
        v24 = self.vedit.lineEdit_23.text()
        v25 = self.vedit.lineEdit_24.text()
        vfecha = self.vedit.dateTimeEdit.dateTime().toString("yyyy-MM-dd HH:mm:ss") #fecha
        v3 = self.vedit.lineEdit_3.text() #NroFactura

        
        v7 = self.vedit.lineEdit_4.text() #idProveedores
        v9 = self.vedit.lineEdit_5.text()
        v10 = self.vedit.lineEdit_6.text()
        v13 = self.vedit.lineEdit_19.text() #idUsuario
        v14 = self.vedit.lineEdit_11.text()

        v15 = self.vedit.lineEdit_12.text() #idProd
        v16 = self.vedit.lineEdit_13.text()
        v17 = self.vedit.lineEdit_20.text() #PUnitario
        #v18 = self.vedit.lineEdit_14.text() #IVA
        v19 = self.vedit.lineEdit_15.text() #Cantidad
        v20 = self.vedit.lineEdit_21.text() #SubtotalProd
        v21 = self.vedit.lineEdit_16.text() #Subtotal
        v22 = self.vedit.lineEdit_17.text() #IVAFactura
        v23 = self.vedit.lineEdit_18.text() #Total
        v26 = self.vedit.lineEdit_25.text() #IVA5%
        

        vfila = self.filaamodificar
        if len(v0) > 0 and len(v3) > 0 and len(v7) > 0 and len(v13) > 0 and len(v23) > 0:
            if v0.isnumeric():
                
                    celda0 = QtWidgets.QTableWidgetItem(v0) #Inserta las filas y columnas despues de validar
                    celda3 = QtWidgets.QTableWidgetItem(v3) 
                    celda4 = QtWidgets.QTableWidgetItem(vfecha)#fecha
                    
                    celda7 = QtWidgets.QTableWidgetItem(v9)
                    celda9 = QtWidgets.QTableWidgetItem(v10) 
                    celda10 = QtWidgets.QTableWidgetItem(v7)
                    celda13 = QtWidgets.QTableWidgetItem(v13)
                    celda14 = QtWidgets.QTableWidgetItem(v14) #

                    try:
                        cursor = self.conex.cursor() 
                        if self.editando: #Modificando
                            #cursor.execute(querysql_update.format(v0, v1, v2, v3, v4, vchk6, v7, v9, self.condicion_update)) #SQL ejecutandose EDITAR
                            pass
                        else: #Creando
                            cursor.execute(querysql_insert.format(v0, v3, vfecha, v13)) #SQL ejecutandose NUEVO 
                            self.vnav.tableWidget.insertRow(vfila) 
                            self.cantfilas += 1

                            nb_row = self.vedit.tableWidget2.rowCount() #Detalle de Venta
                            for row in range (nb_row):
                                colum1 = self.vedit.tableWidget2.item(row, 0).text()
                                colum2 = self.vedit.tableWidget2.item(row, 2).text()
                                colum3 = self.vedit.tableWidget2.item(row, 3).text()
                                colum4 = self.vedit.tableWidget2.item(row, 4).text()
                                cursor.execute(querysql_insert2.format(v0, colum1, colum2, colum3))
                    
                        self.conex.commit()

                        #Hace que los datos nuevos no se modifiquen desde las grillas
                        celda0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda13.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda14.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

                        self.vnav.tableWidget.setItem(vfila, 0, celda0) #Coloca los datos nuevos -- id
                        self.vnav.tableWidget.setItem(vfila, 1, celda3) #Nro Factura
                        self.vnav.tableWidget.setItem(vfila, 2, celda10) #idProveedores
                        self.vnav.tableWidget.setItem(vfila, 3, celda7)
                        self.vnav.tableWidget.setItem(vfila, 4, celda9)
                        self.vnav.tableWidget.setItem(vfila, 5, celda4) #Fecha
                        self.vnav.tableWidget.setItem(vfila, 6, celda13)
                        self.vnav.tableWidget.setItem(vfila, 7, celda13) #idUsuario
                        self.vnav.tableWidget.setItem(vfila, 8, celda14) 

                        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))

                        if self.cantfilas > 0:
                            now = datetime.now()
                            vvfecha = now.strftime("%d-%m-%Y, %H-%M-%S")
                            filename = "Nota Credito Compra {}.pdf".format(str(vvfecha)) #Nombre del archivo
                            model = self.vedit.tableWidget2.model() #Apuntador al tableWidget

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
                            body {
                                font-family: 'Helvetica Neue', 'Helvetica', Helvetica, Arial, sans-serif;
                                text-align: center;
                                color: #777;
                            }

                            body h1 {
                                font-weight: 300;
                                margin-bottom: 0px;
                                padding-bottom: 0px;
                                color: #000;
                            }

                            body h3 {
                                font-weight: 300;
                                margin-top: 10px;
                                margin-bottom: 20px;
                                font-style: italic;
                                color: #555;
                            }

                            body a {
                                color: #06f;
                            }

                            .invoice-box {
                                max-width: 800px;
                                border: 1px solid #eee;
                                box-shadow: 0 0 10px rgba(0, 0, 0, 0.15);
                                font-size: 16px;
                                line-height: 24px;
                                font-family: 'Helvetica Neue', 'Helvetica', Helvetica, Arial, sans-serif;
                                color: #555;
                            }

                            .invoice-box table {
                                width: 100%;
                                line-height: inherit;
                                text-align: left;
                                border-collapse: collapse;
                            }

                            .invoice-box table td {
                                padding: 5px;
                                vertical-align: top;
                            }

                            .invoice-box table tr td:nth-child(2) {
                                text-align: right;
                            }

                            .invoice-box table tr.top table td {
                                padding-bottom: 20px;
                            }

                            .invoice-box table tr.top table td.title {
                                font-size: 45px;
                                line-height: 45px;
                                color: #333;
                            }

                            .invoice-box table tr.information table td {
                                padding-bottom: 40px;
                            }

                            .invoice-box table tr.heading td {
                                background: #eee;
                                border-bottom: 1px solid #ddd;
                                font-weight: bold;
                            }

                            .invoice-box table tr.details td {
                                padding-bottom: 20px;
                            }

                            .invoice-box table tr.item td {
                                border-bottom: 1px solid #eee;
                            }

                            .invoice-box table tr.item.last td {
                                border-bottom: none;
                            }

                            .invoice-box table tr.total td:nth-child(2) {
                                border-top: 2px solid #eee;
                                font-weight: bold;
                            }

                            @media only screen and (max-width: 600px) {
                                .invoice-box table tr.top table td {
                                    width: 100%;
                                    display: block;
                                    text-align: center;
                                }

                                .invoice-box table tr.information table td {
                                    width: 100%;
                                    display: block;
                                    text-align: center;
                                }
                            }
                            
                            </style>
                            </head>"""
                            html += """
                            <h1 id="titulo">{}</h1>
                            """.format(v_titulo)
                            html += """
                                    <body>
                            <div class="invoice-box">
                                <table>
                                    <tr class="top">
                                        <td colspan="2">
                                            <table>
                                                <tr>
                                                    <td class="title">
                                                        <h1>TECH WORLD</h1>
                                                    </td>

                                                    <td> 
                                                        RUC: 4029146-3 <br />
                                                        Nro Nota de Credito: {}<br />
                                                        Factura #{}<br />
                                                        Fecha de Emisión: {}
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>

                                    <tr class="information">
                                        <td colspan="2">
                                            <table>
                                                <tr>
                                                    <td>
                                                        Proveedor Nro: {}<br />
                                                        Razón Social: {}<br />
                                                        RUC/CI: {}
                                                    </td>

                                                </tr>
                                            </table>
                                        </td>
                                    </tr>

                                    <tr class="heading">
                                        <td><b>Usuario</b></td>
                                    </tr>

                                    <tr class="details">

                                        <td>{}</td>
                                    </tr>
                                </table>
                            </div>
                        </body>
                        <br><br>
                                    """.format(v0, v3, vfecha, v7, v9, v10, v14)
                            html += """<div class="invoice-box">"""
                            html += "<table><thead>" #La tabla para reporte
                            html += "<tr>"
                            for c in range(model.columnCount()-2):
                                html += "<th>{}</th>".format(model.headerData(c, QtCore.Qt.Horizontal))

                            html += "</tr></thead>"
                            html += "<tbody>"
                            for r in range(model.rowCount()):
                                html += "<tr>"
                                for c in range(model.columnCount()-2):
                                    html += "<td>{}</td>".format(model.index(r, c).data() or "")
                                html += "</tr>"
                            html += "</tbody></table>"
                            html += """
                            <br><br>
                            <b>Subtotal: </b>{} Gs. <br />
                            <b>IVA (5%): </b>{} Gs.<br />
                            <b>IVA (10%): </b>{} Gs.<br />
                            <b>Total a Pagar: </b>{} Gs.
                            """.format(v21, v26, v22, v23)
                            doc.setHtml(html)
                            doc.setPageSize(QtCore.QSizeF(printer.pageRect().size()))
                            doc.print_(printer)
                            QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(filename)) #Abre el PDF automatico
                        else:
                            self.vnav.statusbar.showMessage("No hay registros para exportar")

                        self.vedit.close()
                    except:
                        self.vedit.statusbar.showMessage("Error al guardar el registro. Tal vez introdujo valores ya existentes o no válidos")

            else:
                self.vedit.statusbar.showMessage("Ingrese valores numéricos en PrecioCosto y PrecioVenta")
        else:
            self.vedit.statusbar.showMessage("Campos obligatorios vacíos. Por favor, rellene los campos")

    def click_actioncancelar(self):
        self.vedit.close()

    #Eliminar registro seleccionado
    def click_actionEliminar(self):
        try: #Pregunta
            txt_valor0 = self.vnav.tableWidget.selectedIndexes()[0].data()
            txt_valor1 = self.vnav.tableWidget.selectedIndexes()[3].data()
            txt_filaAEliminar = self.vnav.tableWidget.currentRow()
            txtmensaje = "<span style='font-size: 14pt; color: #cd1014;'><b>Estas seguro que desea eliminar?</b></span><span style='font-size: 12pt; color: #333333;'><p>Codigo: <b>{}</b><br>Cliente: <b>{}</b></span>".format(txt_valor0, txt_valor1)
            resp_usu = preguntar(txtmensaje)
            if resp_usu: #Proceso de eliminacion
                try:
                    cursor = self.conex.cursor()
                    cursor.execute(querysql_delete2.format(txt_valor0))
                    cursor.execute(querysql_delete.format(txt_valor0))
                    self.conex.commit()
                    self.vnav.tableWidget.removeRow(txt_filaAEliminar)
                    self.cantfilas -= 1
                    self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))
                except:
                    mensajebox("<html><head/><body><p><span style=' font-size:12pt; font-weight:600; color:#7a2e2e;'>No se puede eliminar este registro.</span></p><p><span style=' font-size:10pt; font-weight:600; color:#414141;'>Es posible que se esté utilizando como referencial en otra tabla.</span></p></body></html>")

        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para eliminar")
    

    def click_actionImprimir(self):
        if self.cantfilas > 0:
            now = datetime.now()
            vvfecha = now.strftime("%d-%m-%Y, %H-%M-%S")
            filename = "Nota Credito Compras {}.pdf".format(str(vvfecha)) #Nombre del archivo
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
            campo8, campo11, campo12 in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(str(campodescrip))
            cel2 = QtWidgets.QTableWidgetItem(str(campo3))
            cel3 = QtWidgets.QTableWidgetItem(campo4)
            cel4 = QtWidgets.QTableWidgetItem(campo5)
            
            cel7 = QtWidgets.QTableWidgetItem(str(campo8))
            cel10 = QtWidgets.QTableWidgetItem(str(campo11))
            cel11 = QtWidgets.QTableWidgetItem(campo12)

            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel11.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3)
            self.vnav.tableWidget.setItem(fila, 4, cel4) 
            self.vnav.tableWidget.setItem(fila, 5, cel7)
            self.vnav.tableWidget.setItem(fila, 6, cel10)
            self.vnav.tableWidget.setItem(fila, 7, cel11)#
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))