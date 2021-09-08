Option Compare Database

Private Sub btnAtras_Click()
DoCmd.Close acForm, "FormaReportes"
 DoCmd.OpenForm "MenuPrincipal"
End Sub
