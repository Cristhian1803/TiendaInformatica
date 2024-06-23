import mysql.connector
from PyQt5 import QtWidgets, uic

v_titulonav = '<html><head/><body><p><span style=" font-size:14pt; font-weight:600; color:#8d8d8d;">{}</span></p></body></html>'
textolabelEditar = '<html><head/><body><p><span style=" font-size:12pt; font-weight:600; color:#6f6f6f;"><b>Editar</b> - {}</span></p></body></html>'
textolabelNuevo = '<html><head/><body><p><span style=" font-size:12pt; font-weight:600; color:#6f6f6f;"><b>Nuevo</b> - {}</span></p></body></html>'
nombreiconoNuevo = "iconos/3592812-document-general-letter-note-office-page-paper_107779.png"
nombreiconoEditar = "iconos/3592869-compose-create-edit-edit-file-office-pencil-writing-creative_107746.png"

colornormalventana = "#fffeed"
colornormalbarrabotones = "#ffffff"

colorpegarventana = "#ffe9ce"
colorpegarbarrabotones = "#ffffff"



def conectarbase():
    vhost = "localhost"
    vusuario = "root"
    vpassword = "123456"
    vbasename = "techworld_tapii"
    vpuerto = "3306"

    try:
        conexion = mysql.connector.connect(
        host = vhost,
        user = vusuario,
        password = vpassword,
        database = vbasename,
        port = vpuerto
        )

    except:
        conexion = False

    return conexion

def preguntar(texto):
    dialogopreguntar = uic.loadUi("pregunta.ui")
    dialogopreguntar.buttonBox.button(QtWidgets.QDialogButtonBox.Yes).setText("Si")
    dialogopreguntar.buttonBox.button(QtWidgets.QDialogButtonBox.No).setText("No")
    dialogopreguntar.lb_mensaje.setText(texto)
    resp = dialogopreguntar.exec()
    if resp:
        return True
    else:
        return False

def mensajebox(texto):
    dialogomensaje = uic.loadUi("advertencia.ui")
    dialogomensaje.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setText("Aceptar")
    dialogomensaje.lb_mensaje.setText(texto)
    resp = dialogomensaje.exec()