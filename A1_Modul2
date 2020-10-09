
Option Explicit

Sub ZeileSpalteSub()

    ' Variabelen definieren
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim r As Integer
    Dim o As Integer
    
    
    Dim vorOEH1 As String
    Dim vorOEH2 As String
    Dim vorPlan As String
    vorOEH1 = ""
    vorOEH2 = ""
    vorPlan = ""
    
    Dim m As Integer
    Dim n As Integer
    
    Dim a As Integer
    Dim b As Integer
    
    Dim a2 As Integer
    Dim b2 As Integer
    
    Dim zaehler1 As Integer
    Dim zaehler2 As Integer
    Dim zaehler3 As Integer
    
    
    Dim ZeileAdd As Integer
    Dim SpalteAdd As Integer
    Dim sumZeile As Integer
    Dim sumSpalte As Integer

    Dim sumZeile2 As Integer
    Dim sumSpalte2 As Integer

    ' ------- Rohdaten in 2D Array einspeisen -----
    
    Worksheets("Rohdaten").Activate
    
    Dim ZeilenSpalten() As String
    a = ActiveSheet.UsedRange.Rows.Count
    b = ActiveSheet.UsedRange.Columns.Count

    ReDim ZeilenSpalten(1 To a, 1 To b)
    
    For i = 1 To a
        For j = 1 To b
    
        ZeilenSpalten(i, j) = ActiveSheet.Cells(i, j).Value
        
    
        Next j
    Next i

    ' ------ Array nach OEH1 sortieren ----------

    
    Dim ZeilenSpaltenOEH() As String

    Dim anzOEH1 As Integer
    anzOEH1 = zaehlenOEH1
    
    Dim anzOEH2 As Integer
    anzOEH2 = zaehlenOEH2

    ' Array[ OEH1, PID(Zeile), Werte(Spalte)]
    ReDim ZeilenSpaltenOEH(1 To anzOEH1, 1 To a, 1 To b)
    
    k = 0
    r = 1
    

    
            For i = 2 To a
            
                If (ZeilenSpalten(i, 11) <> vorOEH1) Then
                    k = k + 1
                    r = 1
                    o = 1

                Else
                    r = r + 1
                    o = 1
                End If
                
                For j = 1 To b
                
                    
                    ZeilenSpaltenOEH(k, r, o) = ZeilenSpalten(i, j)
                    
                    
                    o = o + 1

    
                Next j
                
                vorOEH1 = ZeilenSpalten(i, 11)
            Next i
            
    ' ------ Array OEH1 End ---------
    

    ' ----- Anforderung rüber  -----------
    sumZeile = 1
    sumSpalte = 1
    
    ' Verschiedene OEH1
    For i = 1 To anzOEH1
    
        sumSpalte = 1
        
        'OEH1 ID
        ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, 1, 11)
        sumSpalte = sumSpalte + 1
        
        'OEH1 Beschr
        ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, 1, 12)
        ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.Bold = True
        ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.Size = 14
        
        ' VZK berechnung und Anzeige
        sumSpalte = sumSpalte + 1
        
        'GES Ist-VZK berechnen für OEH1
        Dim VZKGesOEH1 As Double
        VZKGesOEH1 = 0
        
        j = 1
        Do While ZeilenSpaltenOEH(i, j, 1) <> ""
            If ZeilenSpaltenOEH(i, j, 10) <> "" Then

                VZKGesOEH1 = VZKGesOEH1 + CDbl(ZeilenSpaltenOEH(i, j, 10))
                
            End If
        
        'MsgBox VZKGesOEH1
            j = j + 1
        Loop

        'GES Soll-VZK berechnen für OEH1
        Dim VZKSollGesOEH1 As Double
        VZKSollGesOEH2 = 0
        
        j = 1
        Do While ZeilenSpaltenOEH(i, j, 1) <> ""
            If ZeilenSpaltenOEH(i, j, 3) <> "" Then

                VZKSollGesOEH1 = VZKGesOEH1 + CDbl(ZeilenSpaltenOEH(i, j, 3))
                
            End If
        
        'MsgBox VZKGesOEH1
            j = j + 1
        Loop
        
        ' VZKGesOEH1
        ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = "Soll-VZK: " & VZKSollGesOEH1 & " Ist-VZK: " & VZKGesOEH1
        If VZKSollGesOEH1 <> VZKGesOEH1 Then
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.ColorIndex = 3
        End If
            
            
        j = 1
        ' Für jede Person
        Do While ZeilenSpaltenOEH(i, j, 1) <> ""
            
            If ZeilenSpaltenOEH(i, j, 13) <> vorOEH2 Then
                sumZeile = sumZeile + 1
                sumSpalte = 1
                
                'OEH2 ID
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 13)
                sumSpalte = sumSpalte + 1
                
                'OEH2 Beschr
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 14)
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.Bold = True
                
                ' VZK berechnung und Anzeige
                sumSpalte = sumSpalte + 1
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = "VZKGesOEH2"
                
                sumZeile = sumZeile + 2
                sumSpalte = sumSpalte + 2
                
                ' Überschrift VZK und TG
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = "VZK"
                sumSpalte = sumSpalte + 1
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = "TG"
                sumZeile = sumZeile + 1

            End If
            
            sumSpalte = 2
                
                
            If ZeilenSpaltenOEH(i, j, 1) <> vorPlan Then
                
                'Planstellen Nr und Beschr.
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 1)
                sumSpalte = sumSpalte + 1
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 2)
                sumSpalte = sumSpalte + 2
            
                'VZK Wert und TG Wert (Soll)
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 3)
                sumSpalte = sumSpalte + 1
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 16)
            
                sumZeile = sumZeile + 1
                sumSpalte = 2
            
                vorPlan = ZeilenSpaltenOEH(i, j, 1)
                
            Else
                sumZeile = sumZeile - 1
            End If
            
            'Pers Nr und Name + Nachname
            ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 5)
            sumSpalte = sumSpalte + 1
            ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 6) & " " & ZeilenSpaltenOEH(i, j, 7) & " " & ZeilenSpaltenOEH(i, j, 8)
            sumSpalte = sumSpalte + 1
            
            
            
            
            'Abfrage falls kein Stampersonal
            If ZeilenSpaltenOEH(i, j, 17) <> "Stammpersonal" Then
            
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 17)
                
                If ZeilenSpaltenOEH(i, j, 18) <> "" Then
                    ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile + 1, sumSpalte).Value = "zurück: " & ZeilenSpaltenOEH(i, j, 18)
                    'sumZeile = sumZeile + 1
                
                ElseIf ZeilenSpaltenOEH(i, j, 19) <> "" Then
                    ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile + 1, sumSpalte).Value = "bis: " & ZeilenSpaltenOEH(i, j, 19)
                    'sumZeile = sumZeile + 1
                End If
                
            End If
            
            
            sumSpalte = sumSpalte + 1
            
            'VZK Wert(IST)
            ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 10)
            
            If ZeilenSpaltenOEH(i, j, 10) <> ZeilenSpaltenOEH(i, j, 3) Then
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.ColorIndex = 3
            End If
            
            'TG Wert (IST)
            sumSpalte = sumSpalte + 1
            ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpaltenOEH(i, j, 15)
            
            Dim TG1 As Integer
            Dim TG2 As Integer
            Dim TGMid As Integer
            
        
            'TG1 = CInt(Mid(ZeilenSpaltenOEH(i, j, 16), 4, 1))
            'TG2 = CInt(Mid(ZeilenSpaltenOEH(i, j, 16), 11, 1))
            'TGMid = CInt(Mid(ZeilenSpaltenOEH(i, j, 15), 4, 1))

            'If Not TG1 < TGMid < TG2 Then
            '    ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Font.ColorIndex = 3
            'End If
            
            sumZeile = sumZeile + 2
            
            
            
            vorOEH2 = ZeilenSpaltenOEH(i, j, 13)
            j = j + 1

        Loop
        
        sumZeile = sumZeile + 2
        
        
    
        
    Next i
    
    
    Worksheets("Anforderung").Activate

End Sub



Public Function zaehlenOEH1() As Integer
    Dim objDic As Object
    Dim zelle As Range

    Worksheets("Rohdaten").Activate

    Set objDic = CreateObject("Scripting.Dictionary")
  
    With ActiveSheet
        For Each zelle In .Range("K2:K300")
    
        If Not objDic.Exists(zelle.Value) Then
            If IsEmpty(zelle.Value) = False Then
                objDic.Add zelle.Value, zelle.Value
            End If
        End If
  
        Next zelle
    
    End With
    zaehlenOEH1 = objDic.Count
    Set objDic = Nothing

End Function


Public Function zaehlenOEH2() As Integer
    Dim objDic As Object
    Dim zelle As Range

    Worksheets("Rohdaten").Activate

    Set objDic = CreateObject("Scripting.Dictionary")
  
    With ActiveSheet
        For Each zelle In .Range("M2:M300")
    
        If Not objDic.Exists(zelle.Value) Then
            If IsEmpty(zelle.Value) = False Then
                objDic.Add zelle.Value, zelle.Value
            End If
        End If
  
        Next zelle
    
    End With
    zaehlenOEH2 = objDic.Count
    Set objDic = Nothing

End Function




Public Sub leer()

    Do While Cells(m, 1).Value <> ""
    
        SpalteAdd = 0
        
        For n = 1 To b
        
            sumZeile = ZeileAdd + m
            sumSpalte = SpalteAdd + n
            
            ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ZeilenSpalten(m, n)
            
            ' Optional falls man keine 0en haben möchte am Ende
            If ZeilenSpalten(m, n) = 0 Then
                ThisWorkbook.Worksheets("Anforderung").Cells(sumZeile, sumSpalte).Value = ""
            End If
            
            ' Überprüfung ob die 3. Spalte erreicht wurde, falls ja wird das nächste feld verschoben
            If n Mod 3 = 0 Then
                SpalteAdd = SpalteAdd - 3
                ZeileAdd = ZeileAdd + 1
                
            End If
        Next n
        
        'MsgBox "M ist " & m
        m = m + 1

    Loop



End Sub


