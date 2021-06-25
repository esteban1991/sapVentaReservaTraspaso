# sapVentaReservaTraspaso
addon para el traspaso de stock entre dos almacenes

Se levanta requerimiento para control de reservas de stock de productos para los ejecutivos de ventas

Actualmente se efectúan transferencias de stock directas desde bodega principal (CD) hacia bodega de reservas (CD_RSV), pero se pierde trazabilidad de estos ya que no se maneja el vendedor responsable y fechas en las que se deba devolver el producto a CD para disponibilizar su stock.

Se propone de solución un add-on para Sap Bo el que pueda llevar el control de los siguientes hitos:

	-	Registro de la solicitud de stock a reservar, indicando vendedor, fecha de solicitud, fecha de vencimiento, estado, id de solicitud, id de transferencia, comentarios, artículos, cantidad, cliente asociado a la reserva, estado de línea.
	-	Control de aprobaciones, el cual el paso anterior al aprobarlo generará un movimiento de transferencia de stock dentro de Sap.
	-	Control de anulaciones, el cual permitirá devolver el stock reservado a CD por movimiento de transferencia de stock.

