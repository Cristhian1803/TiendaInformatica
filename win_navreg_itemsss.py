import sys
from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtGui import QFont, QIcon, QPixmap, QDesktopServices
from tools import * #mensajebox, preguntar, v_titulonav, textolabelEditar, textolabelNuevo, nombreiconoNuevo, nombreiconoEditar
import xlwt
from datetime import datetime

v_titulo = "Items"

querysql_update = """
 update items set 
    idProductos = {},
    CodigoBarras = nullif(trim('{}'),''),
    NombreProd = '{}',
    DescripCorta = nullif(trim('{}'),''),
    PrecioCosto = {},
    PrecioVenta = {},
    activado = {},
    idCategoriasProd = {},
    idmarcas = nullif(trim('{}'),'')
      where idProductos = {}
"""
querysql_insert = """
 insert into items(
    idProductos,
    CodigoBarras,
    NombreProd,
    DescripCorta,
    PrecioCosto,
    PrecioVenta,
    activado,
    idCategoriasProd,  
    idmarcas)
     values(
     {}, 
     nullif(trim('{}'),''), 
     '{}', 
     nullif(trim('{}'),''),      
     {}, 
     {}, 
     {}, 
     {}, 
     {}                                 
     )
"""
querysql_delete = "delete from items where idProductos = {}"
querysql_select = """
SELECT i.idProductos, 
    i.CodigoBarras,
    i.NombreProd,
    i.DescripCorta,
    i.PrecioCosto,
    i.PrecioVenta,
    i.activado,
    i.idCategoriasProd, c.DescripcionCatProd,
    i.idmarcas, m.NombreMarca
FROM items i 
  inner join categorias c on i.idCategoriasProd = c.idCategoriasProd 
  left join marcas m on i.idmarcas = m.idmarcas
order by i.idProductos
"""
querysql_selectmax = "select ifnull(max(idProductos), 0) + 1 from items"
#querysql_selectmax = None --- Habilitar si no se necesita el autoincrementable de arriba
etiquetascol = ["Código", "Código de Barras", "Nombre", "Descripción", "Precio Costo", "Precio Venta", "Disponible", "CodCategoría", "Categoría", "CodMarca", "Marca"]



class iniciar:
    def __init__(self, conexion, listaapegar = None):
        self.mylistaapegar = listaapegar
        self.conex = conexion
        self.vnav = uic.loadUi("win_navreg_items.ui")
        
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
        self.vnav.actionEliminar.triggered.connect(self.click_actionEliminar)
        self.vnav.actionImprimir.triggered.connect(self.click_actionImprimir)
        self.vnav.tableWidget.activated.connect(self.activated_tableWidget)

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
            self.vedit = uic.loadUi("win_edicion_items.ui")
            self.vedit.label_3.setText(etiquetascol[0] + ": ")
            self.vedit.label_4.setText(etiquetascol[1] + ": ")
            self.vedit.label.setText(etiquetascol[2] + ": ")
            self.vedit.label_5.setText(etiquetascol[3] + ": ")
            self.vedit.label_6.setText(etiquetascol[4] + ": ")
            self.vedit.label_7.setText(etiquetascol[5] + ": ")
            self.vedit.label_8.setText(etiquetascol[6] + ": ")
            self.vedit.label_9.setText("Categoría: ")
            self.vedit.label_10.setText("Marca: ")

            self.vedit.lineEdit_2.setMaxLength(25)
            self.vedit.lineEdit_3.setMaxLength(45)
            self.vedit.lineEdit_4.setMaxLength(100)
            self.vedit.lineEdit_7.setReadOnly(True)
            self.vedit.lineEdit_7_2.setReadOnly(True)
            self.vedit.lineEdit_8.setReadOnly(True)
            self.vedit.lineEdit_8_2.setReadOnly(True)

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

                if vchk6.checkState() == QtCore.Qt.Checked: #Recoge los checkBox para la edicion
                    vchk6 = True
                else:
                    vchk6 = False

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
               
                v1 = ""
                v2 = ""
                v3 = ""
                v4 = ""
                v5 = "" 
                vchk6 = False 
                v7 = ""
                v8 = ""
                v9 = ""
                v10 = "" #
                self.filaamodificar = self.cantfilas

            #Carga las variables en LineEdit
            self.vedit.lineEdit.setText(v0)
            self.vedit.lineEdit_2.setText(v1) 
            self.vedit.lineEdit_3.setText(v2)
            self.vedit.lineEdit_4.setText(v3)
            self.vedit.lineEdit_5.setText(v4)
            self.vedit.lineEdit_6.setText(v5)
            self.vedit.checkBox.setChecked(vchk6) 
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
        v3 = self.vedit.lineEdit_4.text()
        v4 = self.vedit.lineEdit_5.text()
        v5 = self.vedit.lineEdit_6.text()

        if self.vedit.checkBox.isChecked():
            vchk6 = "1"
        else:
            vchk6 = "0"
        
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
                    celda3 = QtWidgets.QTableWidgetItem(v3)
                    celda4 = QtWidgets.QTableWidgetItem(v4)
                    celda5 = QtWidgets.QTableWidgetItem(v5) 
                    celda6 = QtWidgets.QTableWidgetItem()
                    
                    if self.vedit.checkBox.isChecked():
                        celda6.setCheckState(QtCore.Qt.Checked)
                    else:
                        celda6.setCheckState(QtCore.Qt.Unchecked)
                    
                    celda7 = QtWidgets.QTableWidgetItem(v7)
                    celda8 = QtWidgets.QTableWidgetItem(v8) 
                    celda9 = QtWidgets.QTableWidgetItem(v9) 
                    celda10 = QtWidgets.QTableWidgetItem(v10) #

                    try:
                        cursor = self.conex.cursor() 
                        if self.editando: #Modificando
                            cursor.execute(querysql_update.format(v0, v1, v2, v3, v4, v5, vchk6, v7, v9, self.condicion_update)) #SQL ejecutandose EDITAR
                        else: #Creando
                            cursor.execute(querysql_insert.format(v0, v1, v2, v3, v4, v5, vchk6, v7, v9)) #SQL ejecutandose NUEVO 
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
                        celda6.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
                        celda10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)

                        self.vnav.tableWidget.setItem(vfila, 0, celda0) #Coloca los datos nuevos
                        self.vnav.tableWidget.setItem(vfila, 1, celda1)
                        self.vnav.tableWidget.setItem(vfila, 2, celda2)
                        self.vnav.tableWidget.setItem(vfila, 3, celda3)
                        self.vnav.tableWidget.setItem(vfila, 4, celda4)
                        self.vnav.tableWidget.setItem(vfila, 5, celda5) 
                        self.vnav.tableWidget.setItem(vfila, 6, celda6)
                        self.vnav.tableWidget.setItem(vfila, 7, celda7)
                        self.vnav.tableWidget.setItem(vfila, 8, celda8)
                        self.vnav.tableWidget.setItem(vfila, 9, celda9)
                        self.vnav.tableWidget.setItem(vfila, 10, celda10) #

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
                    mensajebox("Este registro no se puede eliminar")

        except:
            self.vnav.statusbar.showMessage("Seleccione un registro para eliminar")
    

    def click_actionImprimir(self):
        if self.cantfilas > 0:
            filename = QtWidgets.QFileDialog.getSaveFileName(self.vnav, 'Exportar a Excel', '', ".xls(*.xls)")
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
                QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(filename[0])) #abre el archivo con la extension asociada
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
        for campoid, campodescrip, campo3, campo4, campo5, campo6, campo7, \
            campo8, campo9, campo10, campo11 in cursor: #
            cel0 = QtWidgets.QTableWidgetItem(str(campoid))
            cel1 = QtWidgets.QTableWidgetItem(campodescrip)
            cel2 = QtWidgets.QTableWidgetItem(campo3)
            cel3 = QtWidgets.QTableWidgetItem(campo4)
            cel4 = QtWidgets.QTableWidgetItem(str(campo5))
            cel5 = QtWidgets.QTableWidgetItem(str(campo6))#

            cel6 = QtWidgets.QTableWidgetItem()
            if campo7 ==0:
                cel6.setCheckState(QtCore.Qt.Unchecked)
            else:
                cel6.setCheckState(QtCore.Qt.Checked)
            
            cel7 = QtWidgets.QTableWidgetItem(str(campo8))
            cel8 = QtWidgets.QTableWidgetItem(campo9)
            cel9 = QtWidgets.QTableWidgetItem(str(campo10))
            cel10 = QtWidgets.QTableWidgetItem(campo11)
            
            #Hace que las filas no sean editables
            cel0.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel1.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel2.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel3.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel4.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) 
            cel5.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel6.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel7.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel8.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel9.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            cel10.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable)
            
            self.vnav.tableWidget.insertRow(fila)
            self.vnav.tableWidget.setItem(fila, 0, cel0)
            self.vnav.tableWidget.setItem(fila, 1, cel1)
            self.vnav.tableWidget.setItem(fila, 2, cel2)
            self.vnav.tableWidget.setItem(fila, 3, cel3)
            self.vnav.tableWidget.setItem(fila, 4, cel4)
            self.vnav.tableWidget.setItem(fila, 5, cel5)
            self.vnav.tableWidget.setItem(fila, 6, cel6) 
            self.vnav.tableWidget.setItem(fila, 7, cel7)
            self.vnav.tableWidget.setItem(fila, 8, cel8)
            self.vnav.tableWidget.setItem(fila, 9, cel9)
            self.vnav.tableWidget.setItem(fila, 10, cel10)#
            fila += 1
        self.cantfilas = cursor.rowcount
        if self.cantfilas == -1: self.cantfilas = 0
        self.vnav.lb_cantregs.setText("{} registros encontrados".format(self.cantfilas))