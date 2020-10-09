Sub FehlerAnalyse()



    
    'Call AutoNrFkt
    Call leerFindEinkauf
    Call leerFindVerkauf
    Call Übersicht
 


    
End Sub

Function leerFindEinkauf()
    
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Einkauf As Worksheet
    Set Einkauf = Worksheets("Einkauf")
    
    Dim Verkauf As Worksheet
    Set Verkauf = Worksheets("Verkauf")
    
    Dim Fehler As Worksheet
    Set Fehler = Worksheets("Fehlerliste")
    
    Dim a As Integer
    Dim b As Integer
    
    a = Einkauf.UsedRange.Rows.Count
    b = Einkauf.UsedRange.Columns.Count
    
    Dim FZeile As Integer
    FZeile = 1
    Dim FSpalte As Integer
    FSpalte = 1
    
    Dim Land As String
    Dim Marke As String
    Dim ungetrMarke As String
    
    Dim DatumForm As Date
    'Dim DatumForm As String
    Dim DatumAForm As String
    Dim Tag As String
    Dim Monat As String
    Dim Jahr As String


    Fehler.Cells(FZeile, FSpalte).Value = "Fehler Einkaufstabelle: "
    Fehler.Cells(FZeile, FSpalte).Font.Bold = True
    FZeile = FZeile + 1
    
    
    'Analyse von Einkauf
    For i = 2 To 140
    
    
    
       
    ' ------- Fehlerbehebung ---------
            
            
            For j = 2 To 13
       
                'leer - pro Zelle
                If Einkauf.Cells(i, j).Value = "" Or IsEmpty(Einkauf.Cells(i, j)) = True Or Verkauf.Cells(i, j).Value = "00.00.0000" Then
                    Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " ist leer für AutoNr: " & Einkauf.Cells(i, 1).Value & "  (Zeile: " & i & " Spalte: " & j & ")"
                    FZeile = FZeile + 1
                    
                    Einkauf.Cells(i, j).Interior.Color = RGB(255, 204, 255)

                End If
            Next j
                  
            If Einkauf.Cells(i, 7).Value > 100000 Or Einkauf.Cells(i, 7).Value < 1999 Then
                Einkauf.Cells(i, 7).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Preis ist zu hoch oder zu niedrig, bitte überprüfen. Preis: " & Einkauf.Cells(i, 7).Value & "  (Zeile: " & i & " Spalte: 7)"
                FZeile = FZeile + 1
            End If
                   
            'Land-Hersteller unpassend zur Automarke -  pro Zeile
            Land = Einkauf.Cells(i, 10).Value
            
            ungetrMarke = Einkauf.Cells(i, 2).Value
            Marke = Trim(Einkauf.Cells(i, 2).Value)
            
            If Marke <> ungetrMarke Then
                Einkauf.Cells(i, 2).Interior.Color = RGB(255, 204, 255)
                
                Einkauf.Cells(i, 2).Value = Marke
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Leerzeichen vor oder Nach der Marke (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "VW" And Not Land = "Deutschland" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke VW (Deutschland) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
         
            If Marke = "Mercedes-Benz" And Not Land = "Deutschland" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & "Ist als Automarke Mercedes-Benz (Deutschland) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Opel" And Not Land = "Deutschland" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Opel (Deutschland) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If

            
            If Marke = "Citroen" And Not Land = "Frankreich" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Citroen (Frankreich) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
                        
            If Marke = "Renault" And Not Land = "Frankreich" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Renault (Frankreich) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Dacia" And Not Land = "Rumänien" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Dacia (Rumänien) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Fiat" And Not Land = "Italien" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)

                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Fiat (Italien) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Hyundai" And Not Land = "Südkorea" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Hyundai (Südkorea) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "KIA" And Not Land = "Südkorea" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke KIA (Südkorea) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Fiat" And Not Land = "Italien" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Fiat (Italien) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            If Marke = "Toyota" And Not Land = "Japan" Then
                Einkauf.Cells(i, 10).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Ist als Automarke Toyota (Japan) angegeben, aber das Hersstellland ist: " & Einkauf.Cells(i, 10).Value & "  (Zeile: " & i & " Spalte: 10)"
                FZeile = FZeile + 1
            End If
            
            ' Auftragsnummer und Filliale passen nicht zusammen
            If Right(Einkauf.Cells(i, 12), 1) <> Left(Einkauf.Cells(i, 13), 1) Then
                Einkauf.Cells(i, 12).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = Einkauf.Cells(1, j).Value & " Auftragsnummer und Filliale passen nicht zusammen. Auftragsnummer:  " & Einkauf.Cells(i, 12).Value & " Filliale:  " & Einkauf.Cells(i, 13).Value & "  (Zeile: " & i & " Spalte: 12)"
                FZeile = FZeile + 1
            End If
            
            'Einkaufspreis
            If Einkauf.Cells(i, 6) > 150000 Then
                Einkauf.Cells(i, 6).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = " Kilometerstand über 150.000, bitte überprüfen. Wert:  " & Einkauf.Cells(i, 6).Value & "  (Zeile: " & i & " Spalte: 12)"
                FZeile = FZeile + 1
            End If
            
            'Einkaufspreis
            If Einkauf.Cells(i, 7) > 80000 Then
                Einkauf.Cells(i, 7).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = " Preis über 150.000, bitte überprüfen. Wert:  " & Einkauf.Cells(i, 7).Value & "  (Zeile: " & i & " Spalte: 12)"
                FZeile = FZeile + 1
            End If
            
            'Baujahr Prüfung
            If Val(Einkauf.Cells(i, 5)) > Val(2019) Then
                Einkauf.Cells(i, 5).Interior.Color = RGB(255, 204, 255)
                
                Fehler.Cells(FZeile, FSpalte).Value = " Herstellungsdatum älter als 2019, bitte überprüfen.  " & "  (Zeile: " & i & " Spalte: 5)"
                FZeile = FZeile + 1
            End If
            
            
            'Datum anpassen
            If Mid(Einkauf.Cells(i, 11), 3, 1) = "+" Then
                Tag = Mid(Einkauf.Cells(i, 11), 7, 2)
                Monat = Mid(Einkauf.Cells(i, 11), 4, 2)
                Jahr = Mid(Einkauf.Cells(i, 11), 1, 2)
                If Jahr = "18" Then
                    Jahr = 2018
                End If
                
                
            ElseIf Einkauf.Cells(i, 11) <> "" Then
                Tag = Mid(Einkauf.Cells(i, 11), 7, 2)
                Monat = Mid(Einkauf.Cells(i, 11), 5, 2)
                Jahr = Mid(Einkauf.Cells(i, 11), 1, 4)
                    
                If IsDate(Tag & "." & Monat & "." & Jahr) Then
                    DatumForm = Tag & "." & Monat & "." & Jahr

                Else
                    Einkauf.Cells(i, 11).Interior.Color = RGB(255, 204, 255)
                    Fehler.Cells(FZeile, FSpalte).Value = " Datum ist unpassend. Datum:  " & Einkauf.Cells(i, 11).Value & "  (Zeile: " & i & " Spalte: 12)"
                    FZeile = FZeile + 1
                    


                    Tag = 15
                    
                End If
            Else
                
                Einkauf.Cells(i, 11).Value = DatumForm
            End If
                DatumForm = Tag & "." & Monat & "." & Jahr
                Einkauf.Cells(i, 11).Value = DatumForm
          
          
            'Fillialen - Soweit nur Koeln - anpassen
            If Einkauf.Cells(i, 13).Value = "Koeln" Then
               Einkauf.Cells(i, 13).Value = "Köln"
            End If
            
  Next i
End Function


Function leerFindVerkauf()

    '

    Dim i As Integer
    Dim j As Integer
    Dim j2 As Integer
    
    Set Einkauf = Worksheets("Einkauf")
    
    Dim Verkauf As Worksheet
    Set Verkauf = Worksheets("Verkauf")
    
    Dim Fehler As Worksheet
    Set Fehler = Worksheets("Fehlerliste")
    
    Dim a1 As Integer
    Dim b1 As Integer
    a1 = Einkauf.UsedRange.Rows.Count
    b1 = Einkauf.UsedRange.Columns.Count
    
    Dim a2 As Integer
    Dim b2 As Integer
    a2 = Verkauf.UsedRange.Rows.Count
    b2 = Verkauf.UsedRange.Columns.Count
    
    Dim FZeile As Integer
    Dim FSpalte As Integer
    
    Dim Tag As Integer
    Dim Monat As Integer
    Dim Jahr As Integer
    Dim DateForm As Date
    
    i = 1
    Do While IsEmpty(Fehler.Cells(i, 1)) = False
        FZeile = i
        i = i + 1
    Loop
    
    FSpalte = 1
    Fehler.Cells(FZeile, FSpalte).Value = "Fehler Verkaufstabelle: "
    Fehler.Cells(FZeile, FSpalte).Font.Bold = True
    FZeile = FZeile + 1
    
    For i = 2 To a2
            For j = 2 To 6
       
                'leer - pro Zelle
                If Verkauf.Cells(i, j).Value = "" Or IsEmpty(Verkauf.Cells(i, j)) Or Verkauf.Cells(i, j).Value = "00.00.0000" Then
                    Fehler.Cells(FZeile, FSpalte).Value = Verkauf.Cells(1, j).Value & " ist leer für AutoNr: " & Verkauf.Cells(i, 1).Value & "  (Zeile: " & i & " Spalte: " & j & ")"
                    FZeile = FZeile + 1
                    
                    Verkauf.Cells(i, j).Interior.Color = RGB(255, 204, 255)

                End If
            Next j
            
            'Suche nach passender EK-ID in beiden Tabellen
            
            For j = 2 To a1
                
                
                If Verkauf.Cells(i, 1).Value = Einkauf.Cells(j, 14).Value Then
                    
                    'Verkaufsdatum ist vor Einkaufsdatum
                    DateForm = Verkauf.Cells(i, 6).Value
                    If (DateForm < Einkauf.Cells(j, 11)) And (DateForm <> "00.00.0000") And (Einkauf.Cells(j, 11) <> "12.12.9999") Then
                        Verkauf.Cells(i, 6).Interior.Color = RGB(255, 204, 255)

                        Fehler.Cells(FZeile, FSpalte).Value = " Verkaufsdatum(" & Verkauf.Cells(i, 6).Value & ") vor Einkaufsdatum(" & Einkauf.Cells(j, 11) & ") für VK-ID: " & Verkauf.Cells(i, 1).Value
                        FZeile = FZeile + 1
                    End If
                    
                    'Verkaufspreis ist vor Einkaufspreis
                    If (Verkauf.Cells(i, 5) < Einkauf.Cells(j, 7)) And (Verkauf.Cells(i, 5) <> "") And (Einkauf.Cells(j, 7) <> "") Then
                        Verkauf.Cells(i, 5).Interior.Color = RGB(255, 204, 255)

                        Fehler.Cells(FZeile, FSpalte).Value = " Verkaufspreis(" & Verkauf.Cells(i, 5).Value & ") kleiner als Einkaufspreis(" & Einkauf.Cells(j, 7) & ") für VK-ID: " & Verkauf.Cells(i, 1).Value
                        FZeile = FZeile + 1
                    End If
                    
                End If
            
              Next j
            
            
        
    Next i
    
    
End Function


Function Übersicht()

    Dim i As Integer
    Dim j As Integer
    
    Dim Einkauf As Worksheet
    Set Einkauf = Worksheets("Einkauf (korrigiert)")
    
    Dim Verkauf1 As Worksheet
    Set Verkauf1 = Worksheets("Verkauf (korrigiert)")
    
    Dim Verkauf As Worksheet
    Set Verkauf = Worksheets("Verkauf")
    
    Dim Fehler As Worksheet
    Set Fehler = Worksheets("Fehlerliste")
    
    Dim Ueber As Worksheet
    Set Ueber = Worksheets("Übersicht")
    
    Dim a As Integer
    Dim b As Integer
    
    
    Dim Tag As Integer
    Dim Monat As Integer
    Dim Jahr As Integer
    'Dim Tag As Integer
    Dim Datum1 As Date
    
    a1 = Verkauf1.UsedRange.Rows.Count
    b1 = Verkauf1.UsedRange.Columns.Count
    a2 = Einkauf.UsedRange.Rows.Count
    b2 = Einkauf.UsedRange.Columns.Count
    
    For i = 2 To a2
    
    
    
            'ID
            Ueber.Cells(i, 1).Value = Einkauf.Cells(i, 14).Value
        
            'Tag Ankauf
            Ueber.Cells(i, 2).Value = Day(Einkauf.Cells(i, 11).Value)
        
            'Monat Ankauf
            Ueber.Cells(i, 3).Value = Month(Einkauf.Cells(i, 11).Value)
        
            'Jahr Ankauf
            Ueber.Cells(i, 4).Value = Year(Einkauf.Cells(i, 11).Value)
        
    
        For j = 2 To a1
                               
            'Default

                               
            'Verkaufsdatum
            If Einkauf.Cells(i, 14).Value = Verkauf1.Cells(j, 1).Value Then
 
                    'Tag Verkauf
                    Ueber.Cells(i, 5).Value = Day(Verkauf1.Cells(j, 6).Value)
                    
                    'Monat Verkauf
                    Ueber.Cells(i, 6).Value = Month(Verkauf1.Cells(j, 6).Value)
                    
                    'Jahr Verkauf
                    Ueber.Cells(i, 7).Value = Year(Verkauf1.Cells(j, 6).Value)
                    
                    'Datum
                    Ueber.Cells(i, 8).Value = Verkauf1.Cells(j, 6).Value
                    
            End If
            
            'Default
            If Ueber.Cells(i, 5).Value = "" Or Ueber.Cells(i, 5).Value = "0" Then
                Ueber.Cells(i, 5).Value = 31
                Ueber.Cells(i, 6).Value = 12
                Ueber.Cells(i, 7).Value = 9999
                Datum1 = 31 & "." & 12 & "." & 9999
                Ueber.Cells(i, 8).Value = Datum1
            End If
            
            
            Ueber.Cells(i, 9).Value = Einkauf.Cells(i, 11).Value
            
            
            
        Next j
   
    Next i


End Function

