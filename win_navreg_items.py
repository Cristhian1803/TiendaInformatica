import sys
from PyQt5 import QtWidgets, uic, QtCore, QtGui, QtPrintSupport
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDesktopServices
from tools import * #mensajebox, preguntar, v_titulonav, textolabelEditar, textolabelNuevo, nombreiconoNuevo, nombreiconoEditar
import xlwt
from datetime import datetime
import json

v_titulo = "Productos y Servicios"

with open('sesion.json', 'r') as f: #Carga el JSON con el usuario
    cadena_json = json.load(f)
    idusuario = cadena_json['id']
    usuario = cadena_json['usuario']

querysql_update = """
 update prodservicio set 
    idProductos = {},
    CodigoBarras = nullif(trim('{}'),''),
    NombreProd = '{}',
    PrecioCosto = {},
    PrecioVenta = {},
    idCategoriasProd = {},
    idmarcas = nullif(trim('{}'),'')
      where idProductos = {}
"""
querysql_insert = """
 insert into prodservicio(
    idProductos,
    CodigoBarras,
    NombreProd,
    PrecioCosto,
    PrecioVenta,
    idCategoriasProd,  
    idmarcas)
     values(
     {}, 
     nullif(trim('{}'),''), 
     '{}',    
     {}, 
     {}, 
     {}, 
     {}                                 
     )
"""
querysql_delete = "delete from prodservicio where idProductos = {}"
querysql_select = """
SELECT i.idProductos, 
    i.CodigoBarras,
    i.NombreProd,
    i.PrecioCosto,
    i.PrecioVenta,
    i.idCategoriasProd, c.DescripcionCatProd,
    i.idmarcas, m.NombreMarca
FROM prodservicio i 
  inner join categorias c on i.idCategoriasProd = c.idCategoriasProd 
  left join marcas m on i.idmarcas = m.idmarcas
order by i.idProductos
"""
querysql_selectServ ="""
SELECT i.idProductos, 
    i.CodigoBarras,
    i.NombreProd,
    i.PrecioCosto,
    i.PrecioVenta,
    i.idCategoriasProd, c.DescripcionCatProd,
    i.idmarcas, m.NombreMarca
FROM prodservicio i 
  inner join categorias c on i.idCategoriasProd = c.idCategoriasProd 
  left join marcas m on i.idmarcas = m.idmarcas
where i.idCategoriasProd = 2
order by i.idProductos
"""
querysql_selectmax = "select ifnull(max(idProductos), 0) + 1 from prodservicio"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba
etiquetascol = ["Código", "Código de Barras", "Nombre", "Precio Costo", "Precio Venta", "CodCategoría", "Categoría", "CodMarca", "Marca"]



class iniciar:
    def __init__(self, conexion, listaapegar = None, servicio = ""):
        self.mylistaapegar = listaapegar
        self.conex = conexion

        cursor = self.conex.cursor() #Valida el permiso
        querysql_permiso = "select valor from usuariopermisos where idusuarios like "+idusuario+" and idpermisos like 6"
        cursor.execute(querysql_permiso)
        self.permiso = cursor.fetchone()
        service = servicio
        self.vnav = uic.loadUi("win_navreg_items.ui")
        
        self.cargardatos(service)

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
        self.vnav.actionEliminar.triggered.connect(self.click_actionEliminar)
        if self.permiso == (0,):
            self.vnav.actionEliminar.setEnabled(False)
        self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
        self.vnav.actionExcel.triggered.connect(self.click_actionExcel)
        self.vnav.tableWidget.activated.connect(self.activated_tableWidget)
        

    def activated_tableWidget(self): #Se activa al darle Enter con una fila seleccionada
        if self.mylistaapegar is None:
            self.click_actionEditar()
        else:
            try:
                v0 = self.vnav.tableWidget.selectedIndexes()[0].data()
                v1 = self.vnav.tableWidget.selectedIndexes()[1].data()
                v2 = self.vnav.tableWidget.selectedIndexes()[2].data()
                v3 = self.vnav.tableWidget.selectedIndexes()[3].data()
                v4 = self.vnav.tableWidget.selectedIndexes()[4].data() #
                self.mylistaapegar[0].setText(v0)
                self.mylistaapegar[1].setText(v1)
                self.mylistaapegar[2].setText(v2)
                self.mylistaapegar[3].setText(v3)
                self.mylistaapegar[4].setText(v4) #     
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

                self.vedit = uic.loadUi("win_edicion_items.ui")
                self.vedit.label_3.setText(etiquetascol[0] + ": ")
                self.vedit.label_4.setText(etiquetascol[1] + ": ")
                self.vedit.label.setText(etiquetascol[2] + ": ")
                self.vedit.label_6.setText(etiquetascol[3] + ": ")
                self.vedit.label_7.setText(etiquetascol[4] + ": ")
                self.vedit.label_9.setText("Categoría: ")
                self.vedit.label_10.setText("Marca: ")

                self.vedit.lineEdit_2.setMaxLength(25)
                self.vedit.lineEdit_3.setMaxLength(45)
                self.vedit.lineEdit_7.setReadOnly(True)
                self.vedit.lineEdit_7_2.setReadOnly(True)
                self.vedit.lineEdit_8.setReadOnly(True)
                self.vedit.lineEdit_8_2.setReadOnly(True)
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
                v4 = self.vnav.tableWidget.selectedIndexes()[3].data()
                v5 = self.vnav.tableWidget.selectedIndexes()[4].data()#
                v7 = self.vnav.tableWidget.selectedIndexes()[5].data()
                v8 = self.vnav.tableWidget.selectedIndexes()[6].data()
                v9 = self.vnav.tableWidget.selectedIndexes()[7].data()
                v10 = self.vnav.tableWidget.selectedIndexes()[8].data()#
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
               
                v1 = ""
                v2 = ""
                v4 = ""
                v5 = "" 
                v7 = ""
                v8 = ""
                v9 = ""
                v10 = "" #
                self.filaamodificar = self.cantfilas

            #Carga las variables en LineEdit
            self.vedit.lineEdit.setText(v0)
            self.vedit.lineEdit_2.setText(v1) 
            self.vedit.lineEdit_3.setText(v2)
            self.vedit.lineEdit_5.setText(v4)
            self.vedit.lineEdit_6.setText(v5)
            self.vedit.lineEdit_7.setText(v7)
            self.vedit.lineEdit_7_2.setText(v8)
            self.vedit.lineEdit_8.setText(v9)
            self.vedit.lineEdit_8_2.setText(v10) #
            self.condicion_update = v0 #Captura el id para colocar el update en el mismo lugar en la BD

            self.vedit.setWindowTitle(tituloventana)
            self.vedit.setWindowIcon(iconoVentana)
            self.vedit.lb_titulo.setText(titulointerno)
            self.vedit.label_2.setPixmap(imagenventana)
            self.vedit.show() #Muestra la ventana

            self.vedit.actionAceptar.triggered.connect(self.click_actionaceptar)
            self.vedit.actionCancelar.triggered.connect(self.click_actioncancelar)

            self.vedit.pushButton.clicked.connect(self.clicked_pushButton_cat)
            self.vedit.pushButton_2.clicked.connect(self.clicked_pushButton_marca)

        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para Editar")

    def clicked_pushButton_marca(self):
        import win_navreg_marcas
        self.listalineedit_pks_mar = [self.vedit.lineEdit_8, self.vedit.lineEdit_8_2]
        self.run_win_navreg_marcas = win_navreg_marcas.iniciar(self.conex, self.listalineedit_pks_mar)

    def clicked_pushButton_cat(self):
        import win_navreg_categorias
        self.listalineedit_pks_cat = [self.vedit.lineEdit_7, self.vedit.lineEdit_7_2]
        self.run_win_navreg_categorias = win_navreg_categorias.iniciar(self.conex, self.listalineedit_pks_cat)

    def click_actionaceptar(self):
        v0 = self.vedit.lineEdit.text() #
        v1 = self.vedit.lineEdit_2.text() #Lee los datos del LineEdit
        v2 = self.vedit.lineEdit_3.text()
        v4 = self.vedit.lineEdit_5.text()
        v5 = self.vedit.lineEdit_6.text()
        v7 = self.vedit.lineEdit_7.text()
        v8 = self.vedit.lineEdit_7_2.text()
        v9 = self.vedit.lineEdit_8.text()
        v10 = self.vedit.lineEdit_8_2.text()

        vfila = self.filaamodificar
        if len(v0) > 0 and len(v2) > 0 and len(v4) > 0 and len(v5) > 0 and len(v7) > 0 and len(v9) > 0:
            if v0.isnumeric():
                PCosto = float(v4)
                PVenta = float(v5)
                validaVenta = PCosto - PVenta
                if validaVenta < 0: #Valida si PVenta es mayor a PCosto
                    celda0 = QtWidgets.QTableWidgetItem(v0) #Inserta las filas y columnas despues de validar
                    celda1 = QtWidgets.QTableWidgetItem(v1)
                    celda2 = QtWidgets.QTableWidgetItem(v2)
                    celda4 = QtWidgets.QTableWidgetItem(v4)
                    celda5 = QtWidgets.QTableWidgetItem(v5) 
                    celda7 = QtWidgets.QTableWidgetItem(v7)
                    celda8 = QtWidgets.QTableWidgetItem(v8) 
                    celda9 = QtWidgets.QTableWidgetItem(v9) 
                    celda10 = QtWidgets.QTableWidgetItem(v10) #

                    try:
                        cursor = self.conex.cursor() 
                        if self.editando: #Modificando
                            cursor.execute(querysql_update.format(v0, v1, v2, v4, v5, v7, v9, self.condicion_update)) #SQL ejecutandose EDITAR
                        else: #Creando
                            cursor.execute(querysql_insert.format(v0, v1, v2, v4, v5, v7, v9)) #SQL ejecutandose NUEVO 
                            self.vnav.tableWidget.insertRow(vfila) #ARRIBA - 
                            self.cantfilas += 1
                    
                        self.conex.commit()

                        #Hace que los datos nuevos no se modifiquen desde las grillas
                        celda0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda5.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

                        self.vnav.tableWidget.setItem(vfila, 0, celda0) #Coloca los datos nuevos
                        self.vnav.tableWidget.setItem(vfila, 1, celda1)
                        self.vnav.tableWidget.setItem(vfila, 2, celda2)
                        self.vnav.tableWidget.setItem(vfila, 3, celda4)
                        self.vnav.tableWidget.setItem(vfila, 4, celda5) 
                        self.vnav.tableWidget.setItem(vfila, 5, celda7)
                        self.vnav.tableWidget.setItem(vfila, 6, celda8)
                        self.vnav.tableWidget.setItem(vfila, 7, celda9)
                        self.vnav.tableWidget.setItem(vfila, 8, celda10) #

                        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))

                        self.vedit.close()
                    except:
                        self.vedit.statusbar.showMessage("Error al guardar el registro. Tal vez introdujo valores ya existentes o no válidos")
                else:
                    self.vedit.statusbar.showMessage("PrecioCosto no puede ser superior al PrecioVenta")
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
            txt_valor1 = self.vnav.tableWidget.selectedIndexes()[2].data()
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
            filename = "Productos y Servicios {}.pdf".format(str(vvfecha)) #Nombre del archivo
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
    def cargardatos(self, service):
        cursor = self.conex.cursor()
        if service == 1:
            cursor.execute(querysql_selectServ)
            print("todo lo que sea el select de servicios")
        else:
            cursor.execute(querysql_select)

        fila = 0
        self.vnav.tableWidget.clearContents()
        self.vnav.tableWidget.setRowCount(0)
        for campoid, campodescrip, campo3, campo5, campo6, \
            campo8, campo9, campo10, campo11 in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(campodescrip)
            cel2 = QtWidgets.QTableWidgetItem(campo3)
            cel4 = QtWidgets.QTableWidgetItem(str(campo5))
            cel5 = QtWidgets.QTableWidgetItem(str(campo6))#
            cel7 = QtWidgets.QTableWidgetItem(str(campo8))
            cel8 = QtWidgets.QTableWidgetItem(campo9)
            cel9 = QtWidgets.QTableWidgetItem(str(campo10))
            cel10 = QtWidgets.QTableWidgetItem(campo11)
            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel5.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel4)
            self.vnav.tableWidget.setItem(fila, 4, cel5)
            self.vnav.tableWidget.setItem(fila, 5, cel7)
            self.vnav.tableWidget.setItem(fila, 6, cel8)
            self.vnav.tableWidget.setItem(fila, 7, cel9)
            self.vnav.tableWidget.setItem(fila, 8, cel10)#
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))