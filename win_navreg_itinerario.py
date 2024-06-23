import sys
from PyQt5 import QtWidgets, uic, QtCore, QtGui, QtPrintSupport
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDesktopServices
from PyQt5.Qt import Qt
from tools import * #mensajebox, preguntar, v_titulonav, textolabelEditar, textolabelNuevo, nombreiconoNuevo, nombreiconoEditar
import xlwt
import json
from datetime import datetime

v_titulo = "Itinerarios"

with open('sesion.json', 'r') as f: #Carga el JSON con el usuario
    cadena_json = json.load(f)
    idusuario = cadena_json['id']
    usuario = cadena_json['usuario']

querysql_update = "update itinerario set iditinerario = {}, RazonSocial = '{}', RUC_CI = '{}', Direccion = nullif(trim('{}'),''), TelefCel = nullif(trim('{}'),''), ClienteInterno = {} where iditinerario = {}"
querysql_insert = "insert into itinerario (iditinerario, idempleados, idservicio, fecha, finalizado) values ({}, {}, {}, '{}', {})"
querysql_delete = "delete from itinerario where iditinerario = {}"
querysql_select = """
select 
    it.iditinerario, 
    it.idempleados, 
    e.nombre,
    e.apellido,
    e.cedula,
    it.idservicio,
    s.idclientes,
    c.razonsocial,
    c.ruc_ci,
    it.fecha, 
    it.finalizado 
    from itinerario it
        inner join empleados e on it.idempleados = e.idempleados
        inner join servicio s on it.idservicio = s.idservicio
        inner join clientes c on s.idClientes = c.idClientes
    order by iditinerario
"""
querysql_select2 = """
select
    nc.idProductos,
    p.NombreProd,
    nc.cantidad,
    nc.precio,
    nc.precio * nc.Cantidad as subtotal
    from detservicio nc
        inner join prodservicio p on nc.idProductos = p.idProductos
	where idservicio = {}
    order by idProductos
"""
querysql_selectmax = "select ifnull(max(iditinerario), 0) + 1 from itinerario"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba
etiquetascol = ["Código", "idEmpleado", "Empleado", "RUC CI Empleado", "idServicio", "Cliente", "RUC CI Cliente", "Fecha Inicio", "Finalizado"]

class iniciar:
    def __init__(self, conexion, listaapegar = None):
        self.mylistaapegar = listaapegar
        self.conex = conexion
        cursor = self.conex.cursor() #Valida el permiso
        querysql_permiso = "select valor from usuariopermisos where idusuarios like "+idusuario+" and idpermisos like 1"
        cursor.execute(querysql_permiso)
        self.permiso = cursor.fetchone()

        self.vnav = uic.loadUi("win_navreg_itinerario.ui")
        
        self.cargardatos()

        self.vnav.tableWidget.setHorizontalHeaderLabels(etiquetascol) #Establece texto en las etiquetas
        self.vnav.tableWidget.resizeColumnsToContents() #Reajusta columnas a su contenido

        if self.mylistaapegar is None:
            self.vnav.setStyleSheet("QMainWindow{background-color: " + colornormalventana +";}")
            self.vnav.setWindowTitle(v_titulo)
            self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo))
            #self.vnav.setWindowModality(Qt.ApplicationModal) -- PONE EN FOCUS LA PANTALLA PARA NO CLICKEAR A OTRA VENTANA
            self.vnav.showMaximized()
            self.vnav.show()
        else:
            self.vnav.setStyleSheet("QMainWindow{background-color: " + colorpegarventana +";}")
            self.vnav.setWindowTitle(v_titulo + " > MODO REFERENCIAL")
            self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo + "(REFERENCIAL)"))
            #self.vnav.setWindowModality(Qt.ApplicationModal)
            self.vnav.showMaximized()
            self.vnav.show()

        self.vnav.actionCerrar.triggered.connect(self.click_actioncerrar)
        self.vnav.actionNuevo.triggered.connect(self.click_actionNuevo)
        self.vnav.actionEditar.triggered.connect(self.click_actionEditar)
        self.vnav.actionEliminar.triggered.connect(self.click_actionEliminar)
        if self.permiso == (0,):
            self.vnav.actionEliminar.setEnabled(False)
        self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
        self.vnav.actionExcel.triggered.connect(self.click_actionExcel)
        self.vnav.tableWidget.activated.connect(self.activated_tableWidget)
        self.vnav.actionEditar.setEnabled(False)
        self.vnav.actionEliminar.setEnabled(False)

        self.vnav.actionImprimir.setEnabled(False)
        self.vnav.actionExcel.setEnabled(False)
        

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
                self.vedit = uic.loadUi("win_edicion_itinerario.ui")
                self.vedit.label_3.setText(etiquetascol[0] + ": ")
                self.vedit.label_4.setText(etiquetascol[2] + ": ")
                self.vedit.label.setText(etiquetascol[4] + ": ")
                self.vedit.label_5.setText(etiquetascol[7] + ": ")
                self.vedit.label_8.setText(etiquetascol[8] + ": ")
                self.vedit.label_6.setText("Detalles")

                #self.vedit.lineEdit_2.setMaxLength(45)
                #self.vedit.lineEdit_3.setMaxLength(20)
                #self.vedit.lineEdit_4.setMaxLength(100)
                #self.vedit.lineEdit_5.setMaxLength(100)
                self.vedit.lineEdit_2.setReadOnly(True)
                self.vedit.lineEdit_6.setReadOnly(True)
                self.vedit.lineEdit_7.setReadOnly(True)
                self.vedit.lineEdit_3.setReadOnly(True)
                self.vedit.lineEdit_8.setReadOnly(True)
                self.vedit.lineEdit_9.setReadOnly(True)
                self.vedit.lineEdit_4.setVisible(False)
                self.vedit.lineEdit_4.setReadOnly(True)
            else:
                self.mensaje = uic.loadUi("advertencia.ui")
                self.mensaje.lb_mensaje.setText("<html><head/><body><p><span style=' font-size:12pt; font-weight:600; color:#900000;'>Usted No tiene acceso a esta ventana</span></p></body></html>")
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
                v4 = self.vnav.tableWidget.selectedIndexes()[4].data()#
                
                self.filaamodificar = self.vnav.tableWidget.currentRow() #Modifica en tiempo real la tabla
                vchk6 = self.vnav.tableWidget.item(self.filaamodificar, 5) #

                if vchk6.checkState() == QtCore.Qt.Checked: #Recoge los checkBox para la edicion
                    vchk6 = True
                else:
                    vchk6 = False
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
               
                v1 = ""
                v2 = ""
                v3 = ""
                v4 = ""
                vchk6 = False #
                self.filaamodificar = self.cantfilas

            #Carga las variables en LineEdit
            self.vedit.lineEdit.setText(v0)
            self.vedit.lineEdit_2.setText(v1) 
            self.vedit.lineEdit_3.setText(v2)
            vfecha = QtCore.QDateTime.currentDateTime()
            self.vedit.dateTimeEdit.setDateTime(vfecha)
            #self.vedit.lineEdit_4.setText(v3)
            #self.vedit.lineEdit_5.setText(v4)
            self.vedit.checkBox.setChecked(vchk6) #
            self.condicion_update = v0 #Captura el id para colocar el update en el mismo lugar en la BD

            self.vedit.setWindowTitle(tituloventana)
            self.vedit.setWindowIcon(iconoVentana)
            self.vedit.lb_titulo.setText(titulointerno)
            self.vedit.label_2.setPixmap(imagenventana)
            if self.mylistaapegar is None:
                #self.vedit.setWindowModality(Qt.ApplicationModal)
                self.vedit.show()
            else:
                #self.vedit.setWindowModality(Qt.ApplicationModal)
                self.vedit.show() #Muestra la ventana

            self.vedit.actionAceptar.triggered.connect(self.click_actionaceptar)
            self.vedit.actionCancelar.triggered.connect(self.click_actioncancelar)

            self.vedit.pushButton_2.clicked.connect(self.clicked_pushButton_2)
            self.vedit.pushButton.clicked.connect(self.clicked_pushButton)
            self.vedit.aggProd.clicked.connect(self.aggProd)
        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para Editar")

    def clicked_pushButton(self):
        import win_navreg_serviciorealizado
        self.listalineedit_pks_emp = [self.vedit.lineEdit_3, self.vedit.lineEdit_8, self.vedit.lineEdit_9]
        self.run_win_navreg_serviciorealizado = win_navreg_serviciorealizado.iniciar(self.conex, self.listalineedit_pks_emp)

    def clicked_pushButton_2(self):
        import win_navreg_empleados
        self.listalineedit_pks_emp = [self.vedit.lineEdit_6, self.vedit.lineEdit_2, self.vedit.lineEdit_7]
        self.run_win_navreg_empleados = win_navreg_empleados.iniciar(self.conex, self.listalineedit_pks_emp)

    def aggProd(self): #Trae los productos del Detalle de Factura
        try:

            v0 = self.vedit.lineEdit_3.text()
            cursor = self.conex.cursor()       
            cursor.execute(querysql_select2.format(v0))

            fila = 0
            self.vedit.tableWidget2.clearContents()
            self.vedit.tableWidget2.setRowCount(0)
            for campoid, campodescrip, campo3, campo4, campo5 in cursor: #
                cel0 = QtWidgets.QTableWidgetItem(str(campoid))
                cel1 = QtWidgets.QTableWidgetItem(campodescrip)
                cel2 = QtWidgets.QTableWidgetItem(str(campo3))
                cel3 = QtWidgets.QTableWidgetItem(str(campo4))
                cel4 = QtWidgets.QTableWidgetItem(str(campo5))

                
                #Hace que las filas no sean editables
                cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
                
                self.vedit.tableWidget2.insertRow(fila)
                self.vedit.tableWidget2.setItem(fila, 0, cel0)
                self.vedit.tableWidget2.setItem(fila, 1, cel1)
                self.vedit.tableWidget2.setItem(fila, 2, cel2)
                self.vedit.tableWidget2.setItem(fila, 3, cel3)
                self.vedit.tableWidget2.setItem(fila, 4, cel4) #
                fila += 1
        except:
            self.vedit.statusbar.showMessage("Ocurrió un Error. Recuerde agregar arriba primeramente el id de Servicio")

    def click_actionaceptar(self):
        v0 = self.vedit.lineEdit.text() #Lee los datos del LineEdit
        v1 = self.vedit.lineEdit_6.text() #id empleado
        v2 = self.vedit.lineEdit_2.text()
        v3 = self.vedit.lineEdit_7.text()
        v4 = self.vedit.lineEdit_3.text() #idservicio
        v5 = self.vedit.lineEdit_8.text()
        v7 = self.vedit.lineEdit_9.text()
        vfecha = self.vedit.dateTimeEdit.dateTime().toString("yyyy-MM-dd HH:mm:ss") #fecha
        if self.vedit.checkBox.isChecked():
            vchk6 = "1"
        else:
            vchk6 = "0"
        vfila = self.filaamodificar
        if len(v0) > 0 and len(v1) > 0 and len(v4) > 0:
            celda0 = QtWidgets.QTableWidgetItem(v0) #Inserta las filas y columnas despues de validar
            celda1 = QtWidgets.QTableWidgetItem(v1)
            celda2 = QtWidgets.QTableWidgetItem(v2)
            celda3 = QtWidgets.QTableWidgetItem(v3)
            celda4 = QtWidgets.QTableWidgetItem(v4)
            celda5 = QtWidgets.QTableWidgetItem(v5)
            celda7 = QtWidgets.QTableWidgetItem(v7)
            celda8 = QtWidgets.QTableWidgetItem(vfecha)
            celda6 = QtWidgets.QTableWidgetItem()#
            if self.vedit.checkBox.isChecked():
                celda6.setCheckState(QtCore.Qt.Checked)
            else:
                celda6.setCheckState(QtCore.Qt.Unchecked)

            try:
                cursor = self.conex.cursor() 
                if self.editando: #Modificando
                    cursor.execute(querysql_update.format(v0, v1, v2, v3, v4, vchk6, self.condicion_update)) #SQL ejecutandose EDITAR
                else: #Creando
                    cursor.execute(querysql_insert.format(v0, v1, v4, vfecha, vchk6)) #SQL ejecutandose NUEVO 
                    self.vnav.tableWidget.insertRow(vfila) #ARRIBA - 
                    self.cantfilas += 1
            
                self.conex.commit()

                #Hace que los datos nuevos no se modifiquen desde las grillas
                celda0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda5.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda6.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

                self.vnav.tableWidget.setItem(vfila, 0, celda0) #Coloca los datos nuevos
                self.vnav.tableWidget.setItem(vfila, 1, celda1)
                self.vnav.tableWidget.setItem(vfila, 2, celda2)
                self.vnav.tableWidget.setItem(vfila, 3, celda3)
                self.vnav.tableWidget.setItem(vfila, 4, celda4)
                self.vnav.tableWidget.setItem(vfila, 5, celda5)
                self.vnav.tableWidget.setItem(vfila, 6, celda7)
                self.vnav.tableWidget.setItem(vfila, 7, celda8)
                self.vnav.tableWidget.setItem(vfila, 8, celda6) #

                self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))

                self.vedit.close()
            except:
                self.vedit.statusbar.showMessage("Error al guardar el registro. Tal vez introdujo valores ya existentes o no válidos")
        else:
            self.vedit.statusbar.showMessage("Campos obligatorios vacíos. Por favor, rellene los campos")

    def click_actioncancelar(self):
        self.vedit.close()

    #Eliminar registro seleccionado
    def click_actionEliminar(self):
        try: #Pregunta
            txt_valor0 = self.vnav.tableWidget.selectedIndexes()[0].data()
            txt_valor1 = self.vnav.tableWidget.selectedIndexes()[1].data()
            txt_filaAEliminar = self.vnav.tableWidget.currentRow()
            txtmensaje = "<span style='font-size: 14pt; color: #cd1014;'><b>Estas seguro que desea eliminar?</b></span><span style='font-size: 12pt; color: #333333;'><p>Codigo: <b>{}</b><br>Nombre: <b>{}</b></span>".format(txt_valor0, txt_valor1)
            resp_usu = preguntar(txtmensaje)
            if resp_usu: #Proceso de eliminacion
                try:
                    cursor = self.conex.cursor()
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
            filename = "itinerario {}.pdf".format(str(vvfecha)) #Nombre del archivo
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
        for campoid, campodescrip, campo3, campo4, campo5, campo6, campo7, campo8, campo9, campo10, campo11 in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(str(campodescrip))
            cel2 = QtWidgets.QTableWidgetItem(campo3 + " " + campo4)
            cel3 = QtWidgets.QTableWidgetItem(campo5)
            cel4 = QtWidgets.QTableWidgetItem(str(campo6))
            cel5 = QtWidgets.QTableWidgetItem(campo8)
            cel7 = QtWidgets.QTableWidgetItem(campo9)
            cel8 = QtWidgets.QTableWidgetItem(str(campo10))#

            cel6 = QtWidgets.QTableWidgetItem()
            if campo11 ==0:
                cel6.setCheckState(QtCore.Qt.Unchecked)
            else:
                cel6.setCheckState(QtCore.Qt.Checked)
            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel5.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel6.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3)
            self.vnav.tableWidget.setItem(fila, 4, cel4)
            self.vnav.tableWidget.setItem(fila, 5, cel5)
            self.vnav.tableWidget.setItem(fila, 6, cel7)
            self.vnav.tableWidget.setItem(fila, 7, cel8)
            self.vnav.tableWidget.setItem(fila, 8, cel6) #
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))