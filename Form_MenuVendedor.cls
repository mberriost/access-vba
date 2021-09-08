VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MenuVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCerrar_Click()
DoCmd.Quit
End Sub

Private Sub btnClientes_Click()
DoCmd.Close acForm, "MenuVendedor"
 DoCmd.OpenForm "FormaClientes"
End Sub

Private Sub btnOrden_Click()
DoCmd.Close acForm, "MenuVendedor"
 DoCmd.OpenForm "FormaOrden"
End Sub

