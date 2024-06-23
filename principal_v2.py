import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QFont, QIcon
from tools import conectarbase

class iniciarprincipal:
    def __init__(self):
        #app = QtWidgets.QApplication([])
        self.conex = conectarbase()
        
        if self.conex == False:
            self.mensaje = uic.loadUi("advertencia.ui")
            self.mensaje.lb_mensaje.setText("<html><head/><body><p><span style=' font-size:12pt; font-weight:600; color:#a73636;'>No se pudo establecer conexión con la base de datos.</span></p><p><span style=' font-size:10pt; font-weight:600; color:#3a3a3a;'>El programa se cerrará.</span></p></body></html>") 
            resp = self.mensaje.exec()
            sys.exit()
        else:
            self.v_principal = uic.loadUi("principal_v2.ui")
            self.v_principal.setWindowIcon(QIcon("iconos/4105931-add-to-cart-buy-cart-sell-shop-shopping-cart_113919.png"))
            self.v_principal.showMaximized()
            self.v_principal.show()

            self.v_principal.actionSalir.triggered.connect(self.action_salir)
            self.v_principal.actionMarcas.triggered.connect(self.actionMarcas_abrir)
            self.v_principal.actionCategorias.triggered.connect(self.actionCategorias_abrir)
            self.v_principal.actionProveedores.triggered.connect(self.actionProveedores_abrir)
            self.v_principal.actionClientes.triggered.connect(self.actionClientes_abrir)
            self.v_principal.actionItems.triggered.connect(self.actionItems_abrir)
            self.v_principal.actionVentas.triggered.connect(self.actionVentas_abrir)
            self.v_principal.actionTimbrado.triggered.connect(self.actionTimbrado_abrir)
            self.v_principal.actionUsuarios.triggered.connect(self.actionUsuarios_abrir)
            self.v_principal.actionCompras.triggered.connect(self.actionCompras_abrir)
            self.v_principal.actionCiudad.triggered.connect(self.actionCiudad_abrir)
            self.v_principal.actionTipocobro.triggered.connect(self.actionTipocobro_abrir)
            self.v_principal.actionCargos.triggered.connect(self.actionCargos_abrir)
            self.v_principal.actionEmpleados.triggered.connect(self.actionEmpleados_abrir)
            self.v_principal.actionNota_Credito_Ventas.triggered.connect(self.actionNotaCredVenta_abrir)
            self.v_principal.actionNota_Debito_Ventas.triggered.connect(self.actionNotaDebVenta_abrir)
            self.v_principal.actionNota_Debito_Compras.triggered.connect(self.actionNotaDebCompra_abrir)
            self.v_principal.actionNota_Credito_Compras.triggered.connect(self.actionNotaCredCompra_abrir)
            self.v_principal.actionPresupuesto.triggered.connect(self.actionPresupuesto_abrir)
            self.v_principal.actionOrden_de_Ventas.triggered.connect(self.actionOrdenDeVenta_abrir)
            self.v_principal.actionAperturaCaja.triggered.connect(self.actionAperturaCaja_abrir)
            self.v_principal.actionCierreCaja.triggered.connect(self.actionCierreCaja_abrir)
            self.v_principal.actionArqueoCaja.triggered.connect(self.actionArqueoCaja_abrir)
            self.v_principal.actionCuentasPagar.triggered.connect(self.actionCuentasPagar_abrir)
            self.v_principal.actionCuentasCobrar.triggered.connect(self.actionCuentasCobrar_abrir)
            self.v_principal.actionCajas.triggered.connect(self.actionCajas_abrir)
            self.v_principal.actionItinerarios.triggered.connect(self.actionItinerarios_abrir)
            self.v_principal.actionOrden_de_Servicio.triggered.connect(self.actionOrdenServicio_abrir)
            self.v_principal.actionPedido_de_Compra.triggered.connect(self.actionPedidoCompra_abrir)
            self.v_principal.actionPedido_de_Venta.triggered.connect(self.actionPedidoVenta_abrir)
            self.v_principal.actionSolicitud_de_Servicio.triggered.connect(self.actionSolicitudServicio_abrir)
            self.v_principal.actionServicio_Realizado.triggered.connect(self.actionServicioRealizado_abrir)
            self.v_principal.actionPagos.triggered.connect(self.actionPagos_abrir)
            self.v_principal.actionCobros.triggered.connect(self.actionCobros_abrir)
            #app.exec()

    def actionCobros_abrir(self):
        import win_navreg_cobros
        self.run_win_navreg_cobros = win_navreg_cobros.iniciar(self.conex)

    def actionPagos_abrir(self):
        import win_navreg_pagos
        self.run_win_navreg_pagos = win_navreg_pagos.iniciar(self.conex)

    def actionServicioRealizado_abrir(self):
        import win_navreg_serviciorealizado
        self.run_win_navreg_serviciorealizado = win_navreg_serviciorealizado.iniciar(self.conex)

    def actionSolicitudServicio_abrir(self):
        import win_navreg_solicitudservicio
        self.run_win_navreg_solicitudservicio = win_navreg_solicitudservicio.iniciar(self.conex)

    def actionPedidoVenta_abrir(self):
        import win_navreg_pedidoventa
        self.run_win_navreg_pedidoventa = win_navreg_pedidoventa.iniciar(self.conex)

    def actionPedidoCompra_abrir(self):
        import win_navreg_pedidocompra
        self.run_win_navreg_pedidocompra = win_navreg_pedidocompra.iniciar(self.conex)

    def actionOrdenServicio_abrir(self):
        import win_navreg_ordenservicio
        self.run_win_navreg_ordenservicio = win_navreg_ordenservicio.iniciar(self.conex)

    def actionItinerarios_abrir(self):
        import win_navreg_itinerario
        self.run_win_navreg_itinerario = win_navreg_itinerario.iniciar(self.conex)

    def actionCajas_abrir(self):
        import win_navreg_caja
        self.run_win_navreg_caja = win_navreg_caja.iniciar(self.conex)

    def actionCuentasCobrar_abrir(self):
        import win_navreg_cuentascobrar
        self.run_win_navreg_cuentascobrar = win_navreg_cuentascobrar.iniciar(self.conex)

    def actionCuentasPagar_abrir(self):
        import win_navreg_cuentaspagar
        self.run_win_navreg_cuentaspagar = win_navreg_cuentaspagar.iniciar(self.conex)

    def actionArqueoCaja_abrir(self):
        import win_navreg_arqueo
        self.run_win_navreg_arqueo = win_navreg_arqueo.iniciar(self.conex)

    def actionCierreCaja_abrir(self):
        import win_navreg_cierrecaja
        self.run_win_navreg_cierrecaja = win_navreg_cierrecaja.iniciar(self.conex)

    def actionAperturaCaja_abrir(self):
        import win_navreg_aperturacaja
        self.run_win_navreg_aperturacaja = win_navreg_aperturacaja.iniciar(self.conex)

    def actionOrdenDeVenta_abrir(self):
        import win_navreg_ordenventa
        self.run_win_navreg_ordenventa = win_navreg_ordenventa.iniciar(self.conex)

    def actionPresupuesto_abrir(self):
        import win_navreg_presupuesto
        self.run_win_navreg_presupuesto = win_navreg_presupuesto.iniciar(self.conex)

    def actionNotaCredCompra_abrir(self):
        import win_navreg_notacreditocompras
        self.run_win_navreg_notacreditocompras = win_navreg_notacreditocompras.iniciar(self.conex)

    def actionNotaDebCompra_abrir(self):
        import win_navreg_notadebitocompras
        self.run_win_navreg_notadebitocompras = win_navreg_notadebitocompras.iniciar(self.conex)

    def actionNotaDebVenta_abrir(self):
        import win_navreg_notadebitoventas
        self.run_win_navreg_notadebitoventas = win_navreg_notadebitoventas.iniciar(self.conex)
    
    def actionNotaCredVenta_abrir(self):
        import win_navreg_notacreditoventas
        self.run_win_navreg_notacreditoventas = win_navreg_notacreditoventas.iniciar(self.conex)
    
    def actionEmpleados_abrir(self):
        import win_navreg_empleados
        self.run_win_navreg_empleados = win_navreg_empleados.iniciar(self.conex)

    def actionCargos_abrir(self):
        import win_navreg_cargos
        self.run_win_navreg_cargos = win_navreg_cargos.iniciar(self.conex)

    def actionTipocobro_abrir(self):
        import win_navreg_tipoCobro
        self.run_win_navreg_tipocobro = win_navreg_tipoCobro.iniciar(self.conex)

    def actionCiudad_abrir(self):
        import win_navreg_ciudad
        self.run_win_navreg_ciudad = win_navreg_ciudad.iniciar(self.conex)

    def actionCompras_abrir(self):
        import win_navreg_compras
        self.run_win_navreg_compras = win_navreg_compras.iniciar(self.conex)

    def actionUsuarios_abrir(self):
        import win_navreg_usuarios
        self.run_win_navreg_usuarios = win_navreg_usuarios.iniciar(self.conex)

    def actionTimbrado_abrir(self):
        import win_navreg_timbrado
        self.run_win_navreg_timbrado = win_navreg_timbrado.iniciar(self.conex)

    def actionVentas_abrir(self):
        import win_navreg_ventas
        self.run_win_navreg_ventas = win_navreg_ventas.iniciar(self.conex)

    def actionItems_abrir(self):
        import win_navreg_items
        self.run_win_navreg_items = win_navreg_items.iniciar(self.conex)

    def actionClientes_abrir(self):
        import win_navreg_clientes
        self.run_win_navreg_clientes = win_navreg_clientes.iniciar(self.conex)

    def actionProveedores_abrir(self):
        import win_navreg_proveedores
        self.run_win_navreg_proveedores = win_navreg_proveedores.iniciar(self.conex)

    def actionCategorias_abrir(self):
        import win_navreg_categorias
        self.run_win_navreg_categorias = win_navreg_categorias.iniciar(self.conex)

    def actionMarcas_abrir(self):
        import win_navreg_marcas
        self.run_win_navreg_marcas = win_navreg_marcas.iniciar(self.conex)
    

    def action_salir(self):
        sys.exit()

iniciarprincipal()