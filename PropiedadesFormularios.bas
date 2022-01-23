Attribute VB_Name = "PropiedadesFormularios"
Public Sub CenterMe(frmForm As Form, Ancho As Integer, Alto As Integer)
    With frmForm
        .Width = Ancho
        .Height = Alto
    End With
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub
