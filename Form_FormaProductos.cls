VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
Me.txtID.Value = Nz(DMax("IdProducto", "TblProductos"), 0) + 1
End Sub
'Limpiar el label y el texto de cliente
Private Sub limpiarProducto()
 Me.txtCantidad = ""
 Me.txtDescripcion = ""
 Me.txtPrecio = ""
 Me.txtProducto = ""
 Me.txtID = ""
 End Sub

Private Sub btnRegistrar_Click()
'Acceder a la tabla de Productos
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblProductos", dbOpenDynaset)

'Variables
Dim productoNuevo As String
Dim descripcionNueva As String
Dim cantidadNuevo As String
Dim precioNuevo As String
Dim idNuevo As String

'Verificar que no estén vacíos
If Me.txtID <> "" And Me.txtProducto <> "" And Me.txtDescripcion <> "" And Me.txtCantidad <> "" And Me.txtPrecio <> "" Then
'Verificar que el ID ya no existe
 myRs.FindFirst "IDProducto=" & Me.txtID
  If myRs.NoMatch Then
 
 'Leer lo que el usuario escribio y guardarlo en la variable
 productoNuevo = txtProducto
 cantidadNuevo = txtCantidad
 precioNuevo = txtPrecio
 idNuevo = txtID
 descripcionNueva = txtDescripcion
 'Pasar el contenido de la variable a los campos de la tabla.
 myRs.AddNew
 myRs!Cantidad = cantidadNuevo
 myRs!NombreProducto = productoNuevo
 myRs!Precio = precioNuevo
 myRs!Descripcion = descripcionNueva
 myRs!IdProducto = idNuevo
 myRs.Update
 MsgBox ("Producto registrado.")
 limpiarProducto
 Me.txtID.Value = Nz(DMax("IdProducto", "TblProductos"), 0) + 1
 Else
 MsgBox ("ID ya existe.")
 End If
 Else
 MsgBox ("Los campos no pueden estar vacíos.")
 End If
End Sub


