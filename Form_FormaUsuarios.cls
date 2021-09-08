VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Load()
Me.txtID.Value = Nz(DMax("IdUsuario", "TblUsuarios"), 0) + 1
End Sub
'Limpiar el label y el texto de cliente
Private Sub limpiarUsuario()
Me.txtCategoria = ""
Me.txtContrasena = ""
Me.txtID = ""
Me.txtInicial = ""
Me.txtMaterno = ""
Me.txtNombre = ""
Me.txtPaterno = ""
Me.txtUsuario = ""
End Sub

Private Sub btnRegistrar_Click()
Set myDatabase = CurrentDb()
'Acceder a la tabla Usuarios
Set myRs = myDatabase.OpenRecordset("TblUsuarios", dbOpenDynaset)

'Variables
Dim idNuevo As String
Dim nombreNuevo As String
Dim inicialNueva As String
Dim paternoNuevo As String
Dim maternoNuevo As String
Dim usuarioNuevo As String
Dim contrasenaNueva As String
Dim categoriaNueva As String

'Verificar que los campos no estén vacíos.
If Me.txtPaterno <> "" And Me.txtMaterno <> "" And Me.txtNombre <> "" And Me.txtID <> "" And Me.txtUsuario <> "" And Me.txtContrasena <> "" And Me.txtCategoria <> "" Then
 'Verificar que el ID ya no existe
 myRs.FindFirst "IdUsuario=" & Me.txtID
  If myRs.NoMatch Then
 
 'Leer lo que el usuario escribió y guardarlo en la variable
 idNuevo = txtID
 nombreNuevo = txtNombre
 inicialNueva = txtInicial
 paternoNuevo = txtPaterno
 maternoNuevo = txtMaterno
 usuarioNuevo = txtUsuario
 contrasenaNueva = txtContrasena
 categoriaNueva = txtCategoria
 'Pasar lo de la variable en los campos del usuario
 myRs.AddNew
 myRs!IdUsuario = idNuevo
 myRs!Nombre = nombreNuevo
 myRs!Inicial = inicialNueva
 myRs!ApellidoPaterno = paternoNuevo
 myRs!ApellidoMaterno = maternoNuevo
 myRs!Usuario = usuarioNuevo
 myRs!Contrasena = contrasenaNueva
 myRs!Categoria = categoriaNueva
 
 myRs.Update
 MsgBox ("Usuario registrado.")
 limpiarUsuario
 Me.txtID.Value = Nz(DMax("IdUsuario", "TblUsuarios"), 0) + 1
 Else
 MsgBox ("ID ya existe.")
 End If
 Else
 MsgBox ("Los campos no pueden estar vacíos.")
 End If
End Sub
