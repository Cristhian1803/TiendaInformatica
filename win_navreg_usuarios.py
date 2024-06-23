import sys
from PyQt5 import QtWidgets, uic, QtCore, QtGui, QtPrintSupport
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDesktopServices
from PyQt5.Qt import Qt
from tools import * #mensajebox, preguntar, v_titulonav, textolabelEditar, textolabelNuevo, nombreiconoNuevo, nombreiconoEditar
import xlwt
import json
from datetime import datetime

v_titulo = "Usuarios"

with open('sesion.json', 'r') as f: #Carga el JSON con el usuario
    cadena_json = json.load(f)
    idusuario = cadena_json['id']
    usuario = cadena_json['usuario']

querysql_update = "update usuarios set idUsuarios = {}, usuario = '{}', clave = aes_encrypt('{}', '6020081') where idUsuarios = {}"

querysql_update1 = "update usuariopermisos set idUsuarios = {}, idPermisos = {}, valor = {} where idUsuarios = {} and idPermisos = {}"

querysql_insert = "insert into usuarios (idUsuarios, usuario, clave, idempleados) values ({}, '{}', aes_encrypt('{}', '6020081'), {})"

querysql_insert1 = "insert into usuariopermisos (idUsuarios, idPermisos, valor) values ({}, {}, {})"



querysql_delete = "delete from usuarios where idUsuarios = {}"

querysql_select = """select u.idUsuarios, u.usuario, u.idempleados, e.nombre, e.apellido 
from usuarios u inner join empleados e on u.idempleados = e.idempleados 
where idUsuarios not in (select idUsuarios from usuarios where idUsuarios=1) order by idUsuarios"""
querysql_selectmax = "select ifnull(max(idUsuarios), 0) + 1 from usuarios"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba
etiquetascol = ["Código", "Nombre de Usuario", "idEmpleado", "Empleado"]

class iniciar:
    def __init__(self, conexion, listaapegar = None):
        self.mylistaapegar = listaapegar
        self.conex = conexion

        cursor = self.conex.cursor() #Valida el permiso
        querysql_permiso = "select valor from usuariopermisos where idusuarios like "+idusuario+" and idpermisos like 10"
        cursor.execute(querysql_permiso)
        self.permiso = cursor.fetchone()

        if self.permiso == (1,):
            self.vnav = uic.loadUi("win_navreg_usuarios.ui")
            
            self.cargardatos()

            self.vnav.tableWidget.setHorizontalHeaderLabels(etiquetascol) #Establece texto en las etiquetas
            self.vnav.tableWidget.resizeColumnsToContents() #Reajusta columnas a su contenido

            if self.mylistaapegar is None:
                self.vnav.setStyleSheet("QMainWindow{background-color: " + colornormalventana +";}")
                self.vnav.setWindowTitle(v_titulo)
                self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo))
                self.vnav.setWindowModality(Qt.ApplicationModal)
                self.vnav.show()

            else:
                self.vnav.setStyleSheet("QMainWindow{background-color: " + colorpegarventana +";}")
                self.vnav.setWindowTitle(v_titulo + " > MODO REFERENCIAL")
                self.vnav.lb_tituloform.setText(v_titulonav.format(v_titulo + "(REFERENCIAL)"))
                self.vnav.setWindowModality(Qt.ApplicationModal)
                self.vnav.show()

            self.vnav.actionCerrar.triggered.connect(self.click_actioncerrar)
            self.vnav.actionNuevo.triggered.connect(self.click_actionNuevo)
            self.vnav.actionEditar.triggered.connect(self.click_actionEditar)
            self.vnav.actionEliminar.triggered.connect(self.click_actionEliminar)
            self.vnav.actionEliminar.setEnabled(False)
            #self.vnav.actionNuevo.setEnabled(False)
            self.vnav.actionEditar.setEnabled(False)
            self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
            self.vnav.actionExcel.triggered.connect(self.click_actionExcel)
            self.vnav.tableWidget.activated.connect(self.activated_tableWidget)
        else:
            self.mensaje = uic.loadUi("advertencia.ui")
            self.mensaje.lb_mensaje.setText("<html><head/><body><p><span style=' font-size:12pt; font-weight:600; color:#900000;'>Usted No tiene acceso a esta ventana</span></p></body></html>")
            resp = self.mensaje.exec()

    def activated_tableWidget(self): #Se activa al darle Enter con una fila seleccionada
        if self.mylistaapegar is None:
            self.click_actionEditar()
        else:
            try:
                v0 = self.vnav.tableWidget.selectedIndexes()[0].data()
                v1 = self.vnav.tableWidget.selectedIndexes()[1].data() #
                self.mylistaapegar[0].setText(v0)
                self.mylistaapegar[1].setText(v1) #     
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
            self.vedit = uic.loadUi("win_edicion_usuarios.ui")
            self.vedit.label_3.setText(etiquetascol[0] + ": ")
            self.vedit.label_4.setText(etiquetascol[1] + ": ")
            self.vedit.label_15.setText("Empleado: ")
            self.vedit.label.setText("Clave: ")
            self.vedit.label_5.setText("Clientes: ")
            self.vedit.label_6.setText("Proveedores: ")
            self.vedit.label_7.setText("Categorías: ")
            self.vedit.label_8.setText("Marcas: ")
            self.vedit.label_9.setText("Tipos de Facturas: ")
            self.vedit.label_10.setText("Items: ")
            self.vedit.label_11.setText("Factura de Ventas: ")
            self.vedit.label_12.setText("Factura de Compras: ")
            self.vedit.label_13.setText("Timbrado: ")
            self.vedit.label_16.setText("Empleados: ")
            self.vedit.label_17.setText("Ciudad: ")
            self.vedit.label_18.setText("Cargos: ")
            self.vedit.label_19.setText("Tipo Cobro: ")
            self.vedit.label_20.setText("Presupuesto: ")
            self.vedit.label_21.setText("Cuentas a Pagar: ")
            self.vedit.label_22.setText("Nota Credito Compras: ")
            self.vedit.label_23.setText("Nota Debito Compras: ")
            self.vedit.label_24.setText("Caja: ")
            self.vedit.label_25.setText("Arqueo: ")
            self.vedit.label_26.setText("Orden de Ventas: ")
            self.vedit.label_27.setText("Cuentas a Cobrar: ")
            self.vedit.label_28.setText("Nota Credito Ventas: ")
            self.vedit.label_29.setText("Nota Debito Ventas: ")
            if self.editando:
                tituloventana = "Editar - {}".format(v_titulo)
                titulointerno = textolabelEditar.format(v_titulo)
                iconoVentana = QIcon(nombreiconoEditar)
                imagenventana = QPixmap(nombreiconoEditar)
                v0 = self.vnav.tableWidget.selectedIndexes()[0].data()
                v1 = self.vnav.tableWidget.selectedIndexes()[1].data() #
                self.filaamodificar = self.vnav.tableWidget.currentRow() #Modifica en tiempo real la tabla

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
                #
                v1 = "" 
                self.filaamodificar = self.cantfilas

            #Carga las variables en LineEdit
            self.vedit.lineEdit.setText(v0)
            self.vedit.lineEdit_2.setText(v1) #
            self.condicion_update = v0 #Captura el id para colocar el update en el mismo lugar en la BD

            self.vedit.setWindowTitle(tituloventana)
            self.vedit.setWindowIcon(iconoVentana)
            self.vedit.lb_titulo.setText(titulointerno)
            self.vedit.label_2.setPixmap(imagenventana)
            if self.mylistaapegar is None:
                self.vedit.setWindowModality(Qt.ApplicationModal)
                self.vedit.show()
            else:
                self.vedit.setWindowModality(Qt.ApplicationModal)
                self.vedit.show() #Muestra la ventana

            self.vedit.actionAceptar.triggered.connect(self.click_actionaceptar)
            self.vedit.actionCancelar.triggered.connect(self.click_actioncancelar)
            self.vedit.pushButton.clicked.connect(self.clicked_pushButton)
        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para Editar")

    def clicked_pushButton(self):
        import win_navreg_empleados
        self.listalineedit_pks_emp = [self.vedit.lineEdit_5, self.vedit.lineEdit_4, self.vedit.lineEdit_6]
        self.run_win_navreg_empleados = win_navreg_empleados.iniciar(self.conex, self.listalineedit_pks_emp)

    def click_actionaceptar(self):
        v0 = self.vedit.lineEdit.text() #id
        v1 = self.vedit.lineEdit_2.text() #Lee los datos del LineEdit -- Usuario
        v2 = self.vedit.lineEdit_3.text() #Clave
        v26 = self.vedit.lineEdit_5.text() #idEmpleado
        v27 = self.vedit.lineEdit_4.text()
        v28 = self.vedit.lineEdit_6.text()

        if self.vedit.checkBox.isChecked():
            v3 = "1"
        else:
            v3 = "0"

        if self.vedit.checkBox_2.isChecked():
            v4 = "1"
        else:
            v4 = "0"

        if self.vedit.checkBox_3.isChecked():
            v5 = "1"
        else:
            v5 = "0"

        if self.vedit.checkBox_4.isChecked():
            v6 = "1"
        else:
            v6 = "0"

        if self.vedit.checkBox_5.isChecked():
            v7 = "1"
        else:
            v7 = "0"

        if self.vedit.checkBox_6.isChecked():
            v8 = "1"
        else:
            v8 = "0"

        if self.vedit.checkBox_7.isChecked():
            v9 = "1"
        else:
            v9 = "0"

        if self.vedit.checkBox_8.isChecked():
            v10 = "1"
        else:
            v10 = "0"

        if self.vedit.checkBox_9.isChecked():
            v11 = "1"
        else:
            v11 = "0"

        if self.vedit.checkBox_10.isChecked():
            v12 = "1"
        else:
            v12 = "0"

        if self.vedit.checkBox_11.isChecked():
            v13 = "1"
        else:
            v13 = "0"
        
        if self.vedit.checkBox_12.isChecked():
            v14 = "1"
        else:
            v14 = "0"

        if self.vedit.checkBox_13.isChecked():
            v15 = "1"
        else:
            v15 = "0"

        if self.vedit.checkBox_14.isChecked():
            v16 = "1"
        else:
            v16 = "0"

        if self.vedit.checkBox_15.isChecked():
            v17 = "1"
        else:
            v17 = "0"

        if self.vedit.checkBox_16.isChecked():
            v18 = "1"
        else:
            v18 = "0"

        if self.vedit.checkBox_17.isChecked():
            v19 = "1"
        else:
            v19 = "0"

        if self.vedit.checkBox_18.isChecked():
            v20 = "1"
        else:
            v20 = "0"

        if self.vedit.checkBox_19.isChecked():
            v21 = "1"
        else:
            v21 = "0"

        if self.vedit.checkBox_20.isChecked():
            v22 = "1"
        else:
            v22 = "0"

        if self.vedit.checkBox_21.isChecked():
            v23 = "1"
        else:
            v23 = "0"

        if self.vedit.checkBox_22.isChecked():
            v24 = "1"
        else:
            v24 = "0"

        if self.vedit.checkBox_23.isChecked():
            v25 = "1"
        else:
            v25 = "0"

        vfila = self.filaamodificar
        if len(v0) > 0 and len(v1) > 0 and len(v2) > 3 and len(v26) > 0:
            celda0 = QtWidgets.QTableWidgetItem(v0) #Inserta las filas y columnas despues de validar
            celda1 = QtWidgets.QTableWidgetItem(v1)
            celda2 = QtWidgets.QTableWidgetItem(v26)
            celda3 = QtWidgets.QTableWidgetItem(v27) #
            try:
                cursor = self.conex.cursor() 
                if self.editando: #Modificando
                    cursor.execute(querysql_update.format(v0, v1, v2, self.condicion_update)) #SQL ejecutandose EDITAR
                    cursor.execute(querysql_update1.format(v0, 1, v3, self.condicion_update, 1))
                    cursor.execute(querysql_update1.format(v0, 2, v4, self.condicion_update, 2))
                    cursor.execute(querysql_update1.format(v0, 3, v5, self.condicion_update, 3))
                    cursor.execute(querysql_update1.format(v0, 4, v6, self.condicion_update, 4))
                    cursor.execute(querysql_update1.format(v0, 5, v7, self.condicion_update, 5))
                    cursor.execute(querysql_update1.format(v0, 6, v8, self.condicion_update, 6))
                    cursor.execute(querysql_update1.format(v0, 7, v9, self.condicion_update, 7))
                    cursor.execute(querysql_update1.format(v0, 8, v10, self.condicion_update, 8))
                    cursor.execute(querysql_update1.format(v0, 9, v11, self.condicion_update, 9))
                    cursor.execute(querysql_update1.format(v0, 10, 0, self.condicion_update, 10))

                else: #Creando
                    cursor.execute(querysql_insert.format(v0, v1, v2, v26)) #SQL ejecutandose NUEVO - 
                    cursor.execute(querysql_insert1.format(v0, 1, v3))
                    cursor.execute(querysql_insert1.format(v0, 2, v4))
                    cursor.execute(querysql_insert1.format(v0, 3, v5))
                    cursor.execute(querysql_insert1.format(v0, 4, v6))
                    cursor.execute(querysql_insert1.format(v0, 5, v7))
                    cursor.execute(querysql_insert1.format(v0, 6, v8))
                    cursor.execute(querysql_insert1.format(v0, 7, v9))
                    cursor.execute(querysql_insert1.format(v0, 8, v10))
                    cursor.execute(querysql_insert1.format(v0, 9, v11))
                    cursor.execute(querysql_insert1.format(v0, 10, 0))
                    cursor.execute(querysql_insert1.format(v0, 11, v12))
                    cursor.execute(querysql_insert1.format(v0, 12, v13))
                    cursor.execute(querysql_insert1.format(v0, 13, v14))
                    cursor.execute(querysql_insert1.format(v0, 14, v15))
                    cursor.execute(querysql_insert1.format(v0, 15, v16))
                    cursor.execute(querysql_insert1.format(v0, 16, v17))
                    cursor.execute(querysql_insert1.format(v0, 17, v18))
                    cursor.execute(querysql_insert1.format(v0, 18, v19))
                    cursor.execute(querysql_insert1.format(v0, 19, v20))
                    cursor.execute(querysql_insert1.format(v0, 20, v21))
                    cursor.execute(querysql_insert1.format(v0, 21, v22))
                    cursor.execute(querysql_insert1.format(v0, 22, v23))
                    cursor.execute(querysql_insert1.format(v0, 23, v24))
                    cursor.execute(querysql_insert1.format(v0, 24, v25))
                    self.vnav.tableWidget.insertRow(vfila)
                    self.cantfilas += 1
            
                self.conex.commit()

                #Hace que los datos nuevos no se modifiquen desde las grillas
                celda0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                celda3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

                self.vnav.tableWidget.setItem(vfila, 0, celda0) #Coloca los datos nuevos
                self.vnav.tableWidget.setItem(vfila, 1, celda1)
                self.vnav.tableWidget.setItem(vfila, 2, celda2)
                self.vnav.tableWidget.setItem(vfila, 3, celda3)   #

                self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))

                self.vedit.close()
            except:
                self.vedit.statusbar.showMessage("Error al guardar el registro. Tal vez introdujo valores ya existentes o no válidos")
        else:
            self.vedit.statusbar.showMessage("Campos obligatorios vacíos. Por favor, rellene los campos. La contraseña debe tener al menos 4 caracteres")
            
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
            filename = "usuarios {}.pdf".format(str(vvfecha)) #Nombre del archivo
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
        for campoid, campodescrip, campo3, campo4, campo5 in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(campodescrip)
            cel2 = QtWidgets.QTableWidgetItem(str(campo3))
            cel3 = QtWidgets.QTableWidgetItem(campo4 + " " + campo5) #
            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3) #
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))