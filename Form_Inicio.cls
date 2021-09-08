VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Melissa A. Berrios Torres
'Forma Inicio
'Verifica si el usuario y la contraseña son correctos para abrir la forma de Menú Principal y cerrar esta forma.
'Muestra un mensaje de error si son incorrectos.

Option Compare Database

'Botón Iniciar
Private Sub btnIniciar_Click()

'Verificar que el usuario y la contraseña no estén vacíos.
If Me.txtUsuario <> "" And Me.txtContrasena <> "" Then

'Si los campos no están vacíos, verifica si el usuario está en la base de datos.
If (IsNull(DLookup("Usuario", "TblUsuarios", "Usuario = '" & Me.txtUsuario.Value & "'"))) Then
 
'Si no se encuentra el usuario, muestra un mensaje de error.
 MsgBox ("Usuario no registrado.")

'Si se encuentra el usuario, verifica si la contraseña es correcta
Else
If Me.txtContrasena.Value = DLookup("[Contrasena]", "TblUsuarios", "[Usuario]='" & Me.txtUsuario.Value & "'") Then
            
'Si la contraseña es correcta, cierra esta forma y abre la forma Menú Principal
 If DLookup("[Categoria]", "TblUsuarios", "[Usuario]='" & Me.txtUsuario.Value & "'") = "admin" Then
 DoCmd.Close acForm, "Inicio"
 DoCmd.OpenForm "MenuPrincipal"
 
 Else
 DoCmd.Close acForm, "Inicio"
 DoCmd.OpenForm "MenuVendedor"
 End If
 
'Si la contraseña es incorrecta, se muestra un mensaje de error.
Else
MsgBox ("Contraseña incorrecta.")
End If
End If

'Si el usuario o contraseña están vacío, se muestra un mensaje de error.
Else
 MsgBox ("El usuario y contraseña no pueden estar vacios.")
End If
 

End Sub
'Botón Cerrar
Private Sub btnCerrar_Click()
'Para cerrar la forma
DoCmd.Close acForm, "Inicio"
End Sub
