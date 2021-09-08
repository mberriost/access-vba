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

