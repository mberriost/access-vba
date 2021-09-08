Option Compare Database

Private Sub btnCerrar_Click()
'Salir del sistema
DoCmd.Quit
End Sub

Private Sub btnClientes_Click()
DoCmd.Close acForm, "MenuPrincipal"
 DoCmd.OpenForm "FormaClientes"
End Sub

Private Sub btnOrden_Click()
DoCmd.Close acForm, "MenuPrincipal"
 DoCmd.OpenForm "FormaOrden"
End Sub

Private Sub btnProductos_Click()
DoCmd.Close acForm, "MenuPrincipal"
 DoCmd.OpenForm "FormaProductos"
End Sub

Private Sub btnReportes_Click()
DoCmd.Close acForm, "MenuPrincipal"
 DoCmd.OpenForm "FormaReportes"
End Sub

Private Sub btnUsuarios_Click()
DoCmd.Close acForm, "MenuPrincipal"
 DoCmd.OpenForm "FormaUsuarios"
End Sub
