VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    If Not Memoire Is Nothing Then
        If Target.Value = 0 Then
            Memoire.Activate
    '        ActiveCells.Unselect
        Else
            Application.EnableEvents = False ' => d�sactive les �v�nements
            Target.Interior.Color = RGB(255, 108, 0) 'Couleur orange
            Application.EnableEvents = True ' => r�active les �v�nements
        End If
    Else
        Range("A2").Activate
        Set Memoire = ActiveCell
    End If
    Set Memoire = ActiveCell
    
End Sub

Private Sub Mur(ByVal Target As Range)
    Application.EnableEvents = False ' => d�sactive les �v�nements
    
    Application.EnableEvents = True ' => r�active les �v�nements
End Sub
