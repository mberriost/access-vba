VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormaOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Melissa A. Berrios Torres
'Forma Orden

Option Compare Database
'Limpiar el label y el texto de cliente
Private Sub limpiarCliente()
 LblCliente.Caption = ""
 Me.txtIDCliente = ""
 End Sub
 'Limpiar el label y el texto de producto
 Private Sub limpiarProducto()
 LblProducto.Caption = ""
 LblPrecio.Caption = "$0.00"
 Me.txtIDProducto = ""
 End Sub
 'Limpiar el texto de cantidad
 Private Sub limpiarCantidad()
 Me.txtCantidad = ""
 End Sub
 'Limpiar información de la orden
 Private Sub limpiarOrden()
 LblTotal.Caption = "$0.00"
 Me.txtPago = ""
 End Sub

Private Sub btnProducto_Click()
'Acceder a la tabla DetalleOrden
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblDetalleOrden", dbOpenDynaset)

'Variables
Dim idDetalleNuevo As String
Dim idOrdenNueva As String
Dim idProductoNuevo As String
Dim precioNuevo As String
Dim cantidadNueva As String
Dim total As Currency


'Verificar que no estén vacíos
If Me.txtIDProducto <> "" And Me.txtCantidad <> "" And Me.txtIdOrden <> "" And Me.txtDetalle <> "" Then

 'Leer lo que el usuario escribio y guardarlo en la variable
 idDetalleNuevo = txtDetalle
 idOrdenNueva = txtIdOrden
 idProductoNuevo = txtIDProducto
 cantidadNueva = txtCantidad
 precioNuevo = Me.LblPrecio.Caption
 'Pasar el contenido de la variable a los campos de la tabla.
 myRs.AddNew
 myRs!IdDetalleOrden = idDetalleNuevo
 myRs!IdOrden = idOrdenNueva
 myRs!IdProducto = idProductoNuevo
 myRs!Cantidad = cantidadNueva
 myRs!Precio = precioNuevo
 myRs.Update
 MsgBox ("Producto registrado.")
 
 'Limpiar los campos
 limpiarProducto
 limpiarCantidad
 'Calcular el total de la orden
 total = DSum("[Cantidad]*[Precio]", "TblDetalleOrden", _
    "[IdOrden] = " & [txtIdOrden])
 'Mostrar el total en la label
 Me.LblTotal.Caption = Format(total, "Currency")
 'Buscar el proximo ID disponible y los muestra.
 Me.txtDetalle.Value = Nz(DMax("IdDetalleOrden", "TblDetalleOrden"), 0) + 1

 Else
 MsgBox ("Los campos no pueden estar vacíos.")
 End If
 End Sub

Private Sub btnRegistrar_Click()
'Acceder a la tabla Orden.
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblOrden", dbOpenDynaset)

'Variables
Dim idOrdenNueva As String
Dim idClienteNueva As String
Dim fechaNueva As String
Dim pagoNuevo As String

'Verificar que no estén vacíos
If Me.txtIdOrden <> "" And Me.txtIDCliente <> "" And Me.txtFecha <> "" And Me.txtPago <> "" Then
'Verificar que el ID ya no existe
 myRs.FindFirst "IdOrden=" & Me.txtIdOrden
  If myRs.NoMatch Then
  'Leer lo que el usuario escribio y guardarlo en la variable
 idOrdenNueva = txtIdOrden
 idClienteNueva = txtIDCliente
 fechaNueva = txtFecha
 pagoNuevo = txtPago
 
 'Pasar lo de la variable en los campos del usuario
 myRs.AddNew
 myRs!IdOrden = idOrdenNueva
 myRs!IdCliente = idClienteNueva
 myRs!FechaOrden = fechaNueva
 myRs!TipoPago = pagoNuevo
 myRs.Update
 MsgBox ("Orden registrada.")
 'Llamar la rutina para limpiar la información de la orden
 limpiarOrden
 'Buscar el proximo ID disponible y los muestra.
 Me.txtIdOrden.Value = Nz(DMax("IdOrden", "TblOrden"), 0) + 1
 Else
 MsgBox ("ID ya existe.")
 End If
 Else
 MsgBox ("Los campos no pueden estar vacíos.")
 End If
 
End Sub
Private Sub Form_Load()
'Mostrar la fecha actual en el textbox de la fecha
Me.txtFecha = Date
'Buscar el proximo ID disponible y los muestra.
Me.txtIdOrden.Value = Nz(DMax("IdOrden", "TblOrden"), 0) + 1
Me.txtDetalle.Value = Nz(DMax("IdDetalleOrden", "TblDetalleOrden"), 0) + 1
End Sub

Private Sub txtCantidad_LostFocus()
'Acceder a la tabla de productos.
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblProductos", dbOpenDynaset)

'Verificar que el campo de ID producto no esté vacío.
If Me.txtIDProducto <> "" Then
'Si no está vacío, se busca el ID ingresado en la tabla de productos.
myRs.FindFirst "IdProducto=" & Me.txtIDProducto
 'Si no se encuentra el ID en la tabla, se muestra un mensaje de error.
 If myRs.NoMatch Then
 MsgBox ("Producto no existe.")
 'Desactivar el botón Registrar debido al error.
 btnRegistrar.Enabled = False
 'Si se encuentra el ID del producto se verifica la cantidad disponible.
 Else
 If Me.txtCantidad <> "" Then
 'Guardar el valor ingresado en el textbox de cantidad en la variable cant.
 Dim cant As Integer
 cant = txtCantidad
 'Verificar si la cantidad ingresada es mayor que la cantidad disponible del producto.
If cant > myRs("Cantidad").Value Then
'Si la cantidad es mayor, se muestra un mensaje de error.
MsgBox ("Cantidad no disponible.")
'Llamar la rutina para limpiar el textbox de cantidad
limpiarCantidad
'Desactivar el botón Registrar debido al error.
btnRegistrar.Enabled = False
'Si la cantidad es menor, se muestra un mensaje de que está disponible.
Else

'Activar el botón de Registrar.
btnRegistrar.Enabled = True
End If
'Si la cantidad está vacía, se muestra mensaje de error.
Else
MsgBox ("La cantidad no puede estar vacía.")
End If
End If
'Si el ID del producto está vacío, se muestra mensaje de error.
Else
MsgBox ("El ID del producto no puede estar vacío.")
End If
End Sub

Private Sub txtIDCliente_LostFocus()
'Acceder a la tabla de clientes.
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblClientes", dbOpenDynaset)
'Verificar que el textbox de ID cliente no esté vacío
If Me.txtIDCliente <> "" Then
'Si no está vacío, se verifica que le ID existe en la tabla.
myRs.FindFirst "IdCliente=" & Me.txtIDCliente
'Si el ID no exite en la tabla de clientes, se muestra un mensaje de error.
 If myRs.NoMatch Then
 MsgBox ("No existe este cliente.")
 'Llamar la rutina para limpiar el label del nombre del cliente y textbox de ID.
 limpiarCliente
 'Desactivar el botón Registrar debido al error.
 btnRegistrar.Enabled = False
'Si se encuentra el ID del cliente en la tabla, se muestra el nombre del cliente en el label.
Else
Me.LblCliente.Caption = myRs("Nombre").Value
'Activar el botón Registrar.
btnRegistrar.Enabled = True
End If
'Si el ID del cliente está vacío, se muestra un mensaje de error.
Else
 MsgBox ("El ID del cliente no puede estar vacío.")
End If
End Sub

Private Sub txtIDProducto_LostFocus()
'Acceder a la tabla de productos.
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblProductos", dbOpenDynaset)
'Verificar que el ID no esté vacío
If Me.txtIDProducto <> "" Then
'Si no está vacío, se busca el ID ingresado en la tabla de productos.
myRs.FindFirst "IdProducto=" & Me.txtIDProducto
 'Si no se encuentra el ID en la tabla, se muestra un mensaje de error
 If myRs.NoMatch Then
 MsgBox ("No existe este producto.")
 'Llamar la rutina para limpiar los label del producto y precio, y el textbox de ID Producto.
 limpiarProducto
 'Desactivar el botón registrar debido al error.
 btnRegistrar.Enabled = False
Else
'Si se encuentra el ID ingresado en la tabla, se muestra el nombre y el precio del producto en los labels.
Me.LblProducto.Caption = myRs("NombreProducto").Value
Me.LblPrecio.Caption = Format(myRs("Precio").Value, "Currency")
'Activar el botón Registrar.
btnRegistrar.Enabled = True

End If
'Si el ID el producto está vacío, se encuentra un error.
Else
 MsgBox ("El ID del producto no puede estar vacío.")
End If
End Sub
