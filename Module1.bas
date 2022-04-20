Attribute VB_Name = "Module1"
Global Memoire As Range

Sub labyrinthe()
    Dim ligne As Integer
    Dim colonne As Integer
    Dim compteur As Integer
    
    
    Call initialisation
    
    compt = 1
    Do While (Range("B2") <> Range("AD30"))
        aleatoire = Int(Rnd() * 4)
        'Copie colonne
        'paire
        If aleatoire = 0 Then
            colonne = Int(Rnd() * 27) + 2
            Do While (colonne Mod 2 = 0)
                colonne = Int(Rnd() * 27) + 2
            Loop
            ligne = Int(Rnd() * 29) + 2
            Do While (ligne Mod 2 = 0)
                ligne = Int(Rnd() * 29) + 2
            Loop
            If Cells(ligne, colonne).Value <> 0 Then
                If Cells(ligne, colonne).Value <> Cells(ligne, colonne + 2).Value Then
                    If Cells(ligne, colonne).Value > Cells(ligne, colonne + 2).Value Then
                        Cells(ligne, colonne + 2).Value = Cells(ligne, colonne).Value
                        Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne, colonne + 1).Value = Cells(ligne, colonne).Value
                    Else
                        Cells(ligne, colonne).Value = Cells(ligne, colonne + 2).Value
                        Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne, colonne + 1).Value = Cells(ligne, colonne + 2).Value
                    End If
                End If
            End If
            '''
        ElseIf aleatoire = 1 Then
            'impaire
            colonne = Int(Rnd() * 27) + 2
            Do While (colonne Mod 2 = 1)
                colonne = Int(Rnd() * 27) + 2
            Loop
            ligne = Int(Rnd() * 29) + 2
            Do While (ligne Mod 2 = 1)
                ligne = Int(Rnd() * 29) + 2
            Loop
            If Cells(ligne, colonne).Value <> 0 Then
                If Cells(ligne, colonne).Value <> Cells(ligne, colonne + 2).Value Then
                    If Cells(ligne, colonne).Value > Cells(ligne, colonne + 2).Value Then
                        Cells(ligne, colonne + 2).Value = Cells(ligne, colonne).Value
                        Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne, colonne + 1).Value = Cells(ligne, colonne).Value
                    Else
                        Cells(ligne, colonne).Value = Cells(ligne, colonne + 2).Value
                        Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne, colonne + 1).Value = Cells(ligne, colonne + 2).Value
                    End If
                End If
            End If
        ElseIf aleatoire = 2 Then
            'copie ligne
            'Paire
            colonne = Int(Rnd() * 29) + 2
            Do While (colonne Mod 2 = 0)
                colonne = Int(Rnd() * 29) + 2
            Loop
            ligne = Int(Rnd() * 27) + 2
            Do While (ligne Mod 2 = 0)
                ligne = Int(Rnd() * 27) + 2
            Loop
            If Cells(ligne, colonne).Value <> 0 Then
                If Cells(ligne, colonne).Value <> Cells(ligne + 2, colonne).Value Then
                    If Cells(ligne, colonne).Value > Cells(ligne + 2, colonne).Value Then
                        Cells(ligne + 2, colonne).Value = Cells(ligne, colonne).Value
                        Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne + 1, colonne).Value = Cells(ligne, colonne).Value
                    Else
                        Cells(ligne, colonne).Value = Cells(ligne + 2, colonne).Value
                        Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne + 1, colonne).Value = Cells(ligne + 2, colonne).Value
                    End If
                End If
            End If
        ElseIf aleatoire = 3 Then
            'Impaire
            colonne = Int(Rnd() * 29) + 2
            Do While (colonne Mod 2 = 1)
                colonne = Int(Rnd() * 29) + 2
            Loop
            ligne = Int(Rnd() * 27) + 2
            Do While (ligne Mod 2 = 1)
                ligne = Int(Rnd() * 27) + 2
            Loop
            If Cells(ligne, colonne).Value <> 0 Then
                If Cells(ligne + 2, colonne).Value <> Cells(ligne, colonne).Value Then
                    If Cells(ligne, colonne).Value > Cells(ligne + 2, colonne).Value Then
                        Cells(ligne + 2, colonne).Value = Cells(ligne, colonne).Value
                        Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne + 1, colonne).Value = Cells(ligne, colonne).Value
                    Else
                        Cells(ligne, colonne).Value = Cells(ligne + 2, colonne).Value
                        Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
                        Cells(ligne + 1, colonne).Value = Cells(ligne + 2, colonne).Value
                    End If
                End If
            End If
        End If
        '''
        'copie la valeur la plus grande
        For i = 0 To 2
            For Each cellule In Range("B2:AD30")
                If cellule.Value <> 0 Then
                    If cellule.Offset(0, 1) <> 0 Then
                        If cellule.Offset(0, 1) <> cellule Then
                            If cellule.Offset(0, 1) > cellule Then
                                cellule.Value = cellule.Offset(0, 1).Value
                            ElseIf cellule.Offset(0, 1) < cellule Then
                                cellule.Offset(0, 1).Value = cellule.Value
                            End If
                        End If
                    End If
                End If
                
                If cellule.Value <> 0 Then
                    If cellule.Offset(1, 0) <> 0 Then
                        If cellule.Offset(1, 0) <> cellule Then
                            If cellule.Offset(1, 0) > cellule Then
                                cellule.Value = cellule.Offset(1, 0).Value
                            ElseIf cellule.Offset(1, 0) < cellule Then
                                cellule.Offset(1, 0).Value = cellule.Value
                            End If
                        End If
                    End If
                End If
                
                If cellule.Value <> 0 Then
                    If cellule.Offset(0, -1) <> 0 Then
                        If cellule.Offset(0, -1) <> cellule Then
                            If cellule.Offset(0, -1) > cellule Then
                                cellule.Value = cellule.Offset(0, -1).Value
                            ElseIf cellule.Offset(0, -1) < cellule Then
                                cellule.Offset(0, -1).Value = cellule.Value
                            End If
                        End If
                    End If
                End If
                
                If cellule.Value <> 0 Then
                    If cellule.Offset(-1, 0) <> 0 Then
                        If cellule.Offset(-1, 0) <> cellule Then
                            If cellule.Offset(-1, 0) > cellule Then
                                cellule.Value = cellule.Offset(-1, 0).Value
                            ElseIf cellule.Offset(-1, 0) < cellule Then
                                cellule.Offset(-1, 0).Value = cellule.Value
                            End If
                        End If
                    End If
                End If
            Next cellule
        Next i
        
        
        
        compt = compt + 1
    Loop
    Range("A2").Value = 421
    Range("AE30").Value = 421
    Range("A2").Select
End Sub

Sub initialisation()
  Dim ligne As Integer
    Dim collone As Integer
    Dim compteur As Integer
    
    Range("A2").Value = 0
    Range("AE30").Value = 421
    
    compteur = 0
    'Range("B2:AD30") = 1
    Range("A1:AE31").ColumnWidth = 2.14
    Range("A1:AE1").Interior.Color = RGB(0, 0, 0)
    Range("A1:A31").Interior.Color = RGB(0, 0, 0)
    Range("AE1:AE31").Interior.Color = RGB(0, 0, 0)
    Range("A31:AE31").Interior.Color = RGB(0, 0, 0)
    
    Range("A2").Interior.Color = RGB(255, 255, 255)
    Range("AE30").Interior.Color = RGB(255, 255, 255)
    
    Range("A31:AE31").Value = 0
    Range("A3:A31").Value = 0
    Range("A1:AE1").Value = 0
    Range("AE1:AE29").Value = 0
    
    compt = 1
    For Each cellule In Range("B2:AD30")
        If compteur Mod 2 = 0 Then
            cellule.Interior.Color = RGB(255, 255, 255)
            'cellule.Value = Int(Rnd * 10 + 1)
            cellule.Value = compt
            compt = compt + 1
        Else
            cellule.Interior.Color = RGB(0, 0, 0)
            cellule.Value = 0
        End If
        compteur = compteur + 1
    Next cellule
    
    For i = 3 To 29
        If i Mod 2 = 1 Then
            Range("C" & i).Interior.Color = RGB(0, 0, 0)
            Range("C" & i).Value = 0
            Range("E" & i).Interior.Color = RGB(0, 0, 0)
            Range("E" & i).Value = 0
            Range("G" & i).Interior.Color = RGB(0, 0, 0)
            Range("G" & i).Value = 0
            Range("I" & i).Interior.Color = RGB(0, 0, 0)
            Range("I" & i).Value = 0
            Range("K" & i).Interior.Color = RGB(0, 0, 0)
            Range("K" & i).Value = 0
            Range("M" & i).Interior.Color = RGB(0, 0, 0)
            Range("M" & i).Value = 0
            Range("O" & i).Interior.Color = RGB(0, 0, 0)
            Range("O" & i).Value = 0
            Range("Q" & i).Interior.Color = RGB(0, 0, 0)
            Range("Q" & i).Value = 0
            Range("S" & i).Interior.Color = RGB(0, 0, 0)
            Range("S" & i).Value = 0
            Range("U" & i).Interior.Color = RGB(0, 0, 0)
            Range("U" & i).Value = 0
            Range("W" & i).Interior.Color = RGB(0, 0, 0)
            Range("W" & i).Value = 0
            Range("Y" & i).Interior.Color = RGB(0, 0, 0)
            Range("Y" & i).Value = 0
            Range("AA" & i).Interior.Color = RGB(0, 0, 0)
            Range("AA" & i).Value = 0
            Range("AC" & i).Interior.Color = RGB(0, 0, 0)
            Range("AC" & i).Value = 0
        End If
   Next i
   Range("A2").Activate
   Set Memoire = ActiveCell
End Sub

Sub complexe()
    colonne = Int(Rnd() * 27) + 2
    Do While (colonne Mod 2 = 0)
        colonne = Int(Rnd() * 27) + 2
    Loop
    ligne = Int(Rnd() * 29) + 2
    Do While (ligne Mod 2 = 0)
        ligne = Int(Rnd() * 29) + 2
    Loop
    If Cells(ligne, colonne).Value <> 0 Then
        If Cells(ligne, colonne).Value > Cells(ligne, colonne + 2).Value Then
            Cells(ligne, colonne + 2).Value = Cells(ligne, colonne).Value
            Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
            Cells(ligne, colonne + 1).Value = Cells(ligne, colonne).Value
        Else
            Cells(ligne, colonne).Value = Cells(ligne, colonne + 2).Value
            Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
            Cells(ligne, colonne + 1).Value = Cells(ligne, colonne + 2).Value
        End If
    End If
            '''
        
            'impaire
    colonne = Int(Rnd() * 27) + 2
    Do While (colonne Mod 2 = 1)
        colonne = Int(Rnd() * 27) + 2
    Loop
    ligne = Int(Rnd() * 29) + 2
    Do While (ligne Mod 2 = 1)
        ligne = Int(Rnd() * 29) + 2
    Loop
    If Cells(ligne, colonne).Value <> 0 Then
        If Cells(ligne, colonne).Value > Cells(ligne, colonne + 2).Value Then
            Cells(ligne, colonne + 2).Value = Cells(ligne, colonne).Value
            Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
            Cells(ligne, colonne + 1).Value = Cells(ligne, colonne).Value
        Else
            Cells(ligne, colonne).Value = Cells(ligne, colonne + 2).Value
            Cells(ligne, colonne + 1).Interior.Color = RGB(255, 255, 255)
            Cells(ligne, colonne + 1).Value = Cells(ligne, colonne + 2).Value
        End If
    End If
       
            'copie ligne
            'Paire
    colonne = Int(Rnd() * 29) + 2
    Do While (colonne Mod 2 = 0)
        colonne = Int(Rnd() * 29) + 2
    Loop
    ligne = Int(Rnd() * 27) + 2
    Do While (ligne Mod 2 = 0)
        ligne = Int(Rnd() * 27) + 2
    Loop
    If Cells(ligne, colonne).Value <> 0 Then
        If Cells(ligne, colonne).Value > Cells(ligne + 2, colonne).Value Then
            Cells(ligne + 2, colonne).Value = Cells(ligne, colonne).Value
            Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
            Cells(ligne + 1, colonne).Value = Cells(ligne, colonne).Value
        Else
            Cells(ligne, colonne).Value = Cells(ligne + 2, colonne).Value
            Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
            Cells(ligne + 1, colonne).Value = Cells(ligne + 2, colonne).Value
        End If
    End If
        'Impaire
    colonne = Int(Rnd() * 29) + 2
    Do While (colonne Mod 2 = 1)
            colonne = Int(Rnd() * 29) + 2
    Loop
    ligne = Int(Rnd() * 27) + 2
    Do While (ligne Mod 2 = 1)
        ligne = Int(Rnd() * 27) + 2
    Loop
    If Cells(ligne, colonne).Value <> 0 Then
        If Cells(ligne, colonne).Value > Cells(ligne + 2, colonne).Value Then
            Cells(ligne + 2, colonne).Value = Cells(ligne, colonne).Value
            Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
            Cells(ligne + 1, colonne).Value = Cells(ligne, colonne).Value
        Else
            Cells(ligne, colonne).Value = Cells(ligne + 2, colonne).Value
            Cells(ligne + 1, colonne).Interior.Color = RGB(255, 255, 255)
            Cells(ligne + 1, colonne).Value = Cells(ligne + 2, colonne).Value
        End If
    End If
End Sub

Sub effacer()
    For Each cellule In Range("A1:AE31")
        If cellule = 0 Then
            cellule.Interior.Color = RGB(0, 0, 0)
        Else
            cellule.Interior.Color = RGB(255, 255, 255)
        End If
    Next cellule
    Range("A2").Interior.Color = RGB(255, 255, 255)
    Range("AE30").Interior.Color = RGB(255, 255, 255)
End Sub

Sub Noir()
    For Each cellule In Range("A1:AE31")
        cellule.Interior.Color = RGB(0, 0, 0)
    Next cellule
End Sub

Sub Time()
    timerAvant = Timer
'    Application.Wait Now + TimeValue("0:00:02")
    Call labyrinthe
    MsgBox "Temps d'exécution : " & Timer - timerAvant & " secondes."
End Sub

Sub resoudreLabyrinthe()
    Dim x As Integer
    Dim y As Integer
    
    Call remplissage
    
    x = 1
    y = 2
    Do While (Cells(y, x).Value <> 1)
        Cells(y, x).Select
        If Cells(y + 1, x).Value <> 0 Then
            If Cells(y, x).Value > Cells(y + 1, x).Value Then
                y = y + 1
                Cells(y, x).Select
            End If
        End If
        If Cells(y - 1, x).Value <> 0 Then
            If Cells(y, x).Value > Cells(y - 1, x).Value Then
                y = y - 1
                Cells(y, x).Select
            End If
        End If
        If Cells(y, x + 1).Value <> 0 Then
            If Cells(y, x).Value > Cells(y, x + 1).Value Then
                x = x + 1
                Cells(y, x).Select
            End If
        End If
        If Cells(y, x - 1).Value <> 0 Then
            If Cells(y, x).Value > Cells(y, x - 1).Value Then
                x = x - 1
                Cells(y, x).Select
            End If
        End If
        
    Loop
End Sub

Sub effacer_nb()
    For Each cellule In Range("A1:AE31")
        If cellule.Value <> 0 Then
            cellule.Value = ""
        End If
    Next cellule
    
End Sub

Sub remplissage()
    Call effacer_nb
    Cells(30, 31).Value = 1
    Cells(2, 1).Value = "fin"
    Call voisin(30, 30, 1)

End Sub
Sub voisin(x As Integer, y As Integer, N As Integer)
    If Cells(x, y).Value <> "fin" Then
        Cells(x, y).Value = N + 1
        If Cells(x, y - 1).Value = "" Then
            Call voisin(x, y - 1, N + 1)
        End If
        If Cells(x, y + 1).Value = "" Then
            Call voisin(x, y + 1, N + 1)
        End If
        If Cells(x + 1, y).Value = "" Then
            Call voisin(x + 1, y, N + 1)
        End If
        If Cells(x - 1, y).Value = "" Then
            Call voisin(x - 1, y, N + 1)
        End If
    End If

End Sub

