VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'Limpiar los campos
Private Sub limpiarCliente()
 Me.txtID = ""
 Me.txtNombre = ""
 Me.txtCodigo = ""
 Me.txtDireccion = ""
 Me.txtEstado = ""
 Me.txtInicial = ""
 Me.txtNombre = ""
 Me.txtPaterno = ""
 Me.txtPueblo = ""
 Me.txtMaterno = ""
 Me.txtTelefono = ""
 End Sub
 Private Sub Form_Load()
 'Busca el próximo ID y lo muestra
Me.txtID.Value = Nz(DMax("IdCliente", "TblClientes"), 0) + 1
End Sub

Private Sub btnRegistrar_Click()
'Acceder a la tabla Clientes
Set myDatabase = CurrentDb()
Set myRs = myDatabase.OpenRecordset("TblClientes", dbOpenDynaset)

'Variables
Dim idNuevo As String
Dim nombreNuevo As String
Dim inicialNueva As String
Dim paternoNuevo As String
Dim maternoNuevo As String
Dim telefonoNuevo As String
Dim direccionNueva As String
Dim puebloNuevo As String
Dim estadoNuevo As String
Dim codigoNuevo As String
'Verificar que los campos no estén vacíos.
If Me.txtID <> "" And Me.txtNombre <> "" And Me.txtPaterno <> "" And Me.txtMaterno <> "" And Me.txtTelefono <> "" And Me.txtDireccion <> "" And Me.txtPueblo <> "" And Me.txtEstado <> "" And Me.txtCodigo <> "" Then

'Verificar que el ID ya no existe
 myRs.FindFirst "IdCliente=" & Me.txtID
  If myRs.NoMatch Then
  
'Leer lo que el usuario escribio y guardarlo en la variable
 idNuevo = txtID
 nombreNuevo = txtNombre
 inicialNueva = txtInicial
 paternoNuevo = txtPaterno
 maternoNuevo = txtMaterno
 telefonoNuevo = txtTelefono
 direccionNueva = txtDireccion
 puebloNuevo = txtPueblo
 estadoNuevo = txtEstado
 codigoNuevo = txtCodigo
 'Pasar el contenido la variable en los campos de la tabla de clientes.
 myRs.AddNew
 myRs!IdCliente = idNuevo
 myRs!Nombre = nombreNuevo
 myRs!Inicial = inicialNueva
 myRs!ApellidoPaterno = paternoNuevo
 myRs!ApellidoMaterno = maternoNuevo
 myRs!Telefono = telefonoNuevo
 myRs!DireccionPostal1 = direccionNueva
 myRs!Pueblo = puebloNuevo
 myRs!Zipcode = codigoNuevo
 myRs.Update
 MsgBox ("Cliente registrado.")
 limpiarCliente
 Me.txtID.Value = Nz(DMax("IdCliente", "TblClientes"), 0) + 1
 Else
 'Si el ID ya existe, se muestra un mensaje de error.
 MsgBox ("ID ya existe.")
 End If
 Else
 'Si los campos están vacíos, se muestra un mensaje de error.
 MsgBox ("Los campos no pueden estar vacíos.")
 End If
 End Sub



