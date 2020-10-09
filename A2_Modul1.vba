Option Explicit

Sub FaktenTabelleBefuellung()


    fkt_SynonymisierungID
    
    'fkt_LagerbestandSeg
    
    'fkt_ZeitAufLager
    
    'fkt_AutoAlter
    
End Sub

Function fkt_SynonymisierungID()
    
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Lager As Worksheet
    Set Lager = Worksheets("Lager")
    
    Dim LagerOG As Worksheet
    Set LagerOG = Worksheets("Lager(OG)")
    
    Dim a As Integer
    Dim b As Integer
    
    a = LagerOG.UsedRange.Rows.Count
    b = LagerOG.UsedRange.Columns.Count
    
    Dim Tag As String
    Dim Monat As String
    Dim Jahr As String

    Dim land As String
    Dim marke As String
    
    
    'Für Jedes Auto in Lager(OG)
    For i = 2 To a
    
            'Auto ID
            Lager.Cells(i, 1).Value = LagerOG.Cells(i, 1).Value
               
            'Marke ID
            marke = Trim(LagerOG.Cells(i, 2).Value)
            Lager.Cells(i, 2).Value = id_Marke(marke)
            
            'Modell ID
            
            'Farbe ID
            
            'Jahr ID
            Lager.Cells(i, 5).Value = id_Jahr(LagerOG.Cells(i, 5).Value)

            'Land ID
            Lager.Cells(i, 6).Value = id_Land(LagerOG.Cells(i, 10).Value)
            
            'Km ID
            
            'Preissegment ID
            Lager.Cells(i, 8).Value = id_Preissegment(Val(LagerOG.Cells(i, 7).Value))
            
            'Einkaufpreis
            Lager.Cells(i, 9).Value = LagerOG.Cells(i, 7).Value
            
            'PS ID
            
            'Motor ID
            Lager.Cells(i, 11).Value = id_Motor(LagerOG.Cells(i, 9).Value)
  
            'Einkaufs Datum
            Lager.Cells(i, 12).Value = LagerOG.Cells(i, 11).Value
 
            'Verkaufs Datum
            Lager.Cells(i, 13).Value = LagerOG.Cells(i, 12).Value
 
            'Filiale ID
            Lager.Cells(i, 14).Value = id_Filiale(LagerOG.Cells(i, 14).Value)
 
            'Tag ID (Ankauf)
            Lager.Cells(i, 15).Value = id_Tag(LagerOG.Cells(i, 15).Value)
 
            
            'Monat ID (Ankauf)
            Lager.Cells(i, 16).Value = id_Monat(LagerOG.Cells(i, 16).Value)

            'Jahr ID (Ankauf)
            Lager.Cells(i, 17).Value = id_Jahr(LagerOG.Cells(i, 17).Value)

            'Tag ID (Verkauf)
            Lager.Cells(i, 18).Value = id_Tag(LagerOG.Cells(i, 18).Value)
 
            'Monat ID (Verkauf)
            Lager.Cells(i, 19).Value = id_Monat(LagerOG.Cells(i, 19).Value)

            'Jahr (Verkauf)
            Lager.Cells(i, 20).Value = id_Jahr(LagerOG.Cells(i, 20).Value)

            'Tage auf Lager
            Lager.Cells(i, 21).Value = tageAufLager(LagerOG.Cells(i, 12).Value, LagerOG.Cells(i, 11).Value)

  Next i
End Function

Function id_Marke(marke As String) As String
            
            Dim retMarke As String
            
            Select Case marke
                Case Is = "Citroen"
                    retMarke = "MK-01"
                    
                Case Is = "Dacia"
                    retMarke = "MK-02"
                    
                Case Is = "Fiat"
                    retMarke = "MK-03"
                    
                Case Is = "Hyundai"
                    retMarke = "MK-04"
                    
                Case Is = "KIA"
                    retMarke = "MK-05"
                    
                Case Is = "Mercedes-Benz"
                    retMarke = "MK-06"
                    
                Case Is = "Opel"
                    retMarke = "MK-07"
                 
                Case Is = "Renault"
                    retMarke = "MK-08"
                    
                Case Is = "Toyota"
                    retMarke = "MK-09"
                    
                Case Is = "VW"
                    retMarke = "MK-10"
            End Select
            
            id_Marke = retMarke

End Function


Function id_Jahr(Jahr As Integer) As String

            Dim retJahr As String
            
            Select Case Jahr
                Case Is = "2021"
                    retJahr = "J-2021"
                    
                Case Is = "2020"
                    retJahr = "J-2020"
                    
                Case Is = "2019"
                    retJahr = "J-2019"
                    
                Case Is = "2018"
                    retJahr = "J-2018"
                    
                Case Is = "2017"
                    retJahr = "J-2017"
                    
                Case Is = "2016"
                    retJahr = "J-2016"
                    
                Case Is = "2015"
                    retJahr = "J-2015"
                 
                Case Is = "2014"
                    retJahr = "J-2014"
                    
                Case Is = "2013"
                    retJahr = "J-2013"
                    
                Case Is = "2012"
                    retJahr = "J-2012"
                    
                Case Is = "2011"
                    retJahr = "J-1201"
                    
                Case Is = "2010"
                    retJahr = "J-2010"
                    
                Case Is = "2009"
                    retJahr = "J-2009"
                    
                Case Is = "2008"
                    retJahr = "J-2008"
                    
                Case Is = "2007"
                    retJahr = "J-2007"
                    
                Case Is = "2006"
                    retJahr = "J-2006"
                    
                Case Is = "2005"
                    retJahr = "J-2005"
                    
                Case Is = "2004"
                    retJahr = "J-2004"
                 
                Case Is = "2003"
                    retJahr = "J-2003"
                    
                Case Is = "2002"
                    retJahr = "J-2002"
                    
                Case Is = "2001"
                    retJahr = "J-2001"
                    
                Case Is = "2000"
                    retJahr = "J-2000"
                                        
                Case Is = "1999"
                    retJahr = "J-1999"
                 
                Case Is = "1998"
                    retJahr = "J-1998"
                    
                Case Is = "1997"
                    retJahr = "J-1997"
                    
                Case Is = "1996"
                    retJahr = "J-1996"
                    
                Case Is = "1995"
                    retJahr = "J-1995"
                
                Case Is = "9999"
                    retJahr = "J-9999"
                    
            End Select
            
            id_Jahr = retJahr

End Function

Function id_Land(land As String) As String
            
            Dim retLand As String
            
            Select Case land
                Case Is = "Deutschland"
                    retLand = "HLA-001"
                    
                Case Is = "Frankreich"
                    retLand = "HLA-002"
                    
                Case Is = "Japan"
                    retLand = "HLA-003"
                    
                Case Is = "Rumänien"
                    retLand = "HLA-004"
                    
                Case Is = "Südkorea"
                    retLand = "HLA-005"
                    
                Case Is = "Italien"
                    retLand = "HLA-006"
                    
            End Select
            
            id_Land = retLand

End Function



Function id_Preissegment(preisSeg As Double) As String
            
            Dim retPreisSeg As String
            
            Select Case preisSeg
                Case Is <= 9999
                retPreisSeg = "PSGT-01"
                    
                Case Is <= 39999
                    retPreisSeg = "PSGT-02"
                    
                Case Is >= 40000
                    retPreisSeg = "PSGT-03"
                    
            End Select
            
            id_Preissegment = retPreisSeg

End Function


Function id_Motor(motor As String) As String
            
            Dim retMotor As String
            
            Select Case motor
                Case Is = "Benzin"
                    retMotor = "MTR-1"
                    
                Case Is = "Diesel"
                    retMotor = "MTR-2"
                    
            End Select
            
            id_Motor = retMotor

End Function

Function id_Filiale(filiale As String) As String

                    
            Dim retFiliale As String
            
            Select Case filiale
                Case Is = "Dresden"
                    retFiliale = "FL-0001"
                    
                Case Is = "Eisenach"
                    retFiliale = "FL-0002"

                Case Is = "Erfurt"
                    retFiliale = "FL-0003"
            
                Case Is = "Gotha"
                    retFiliale = "FL-0004"

                Case Is = "Halle"
                    retFiliale = "FL-0005"
            
                Case Is = "Karlsruhe"
                    retFiliale = "FL-0006"

                Case Is = "Köln"
                    retFiliale = "FL-0007"
            

                Case Is = "Leipzig"
                    retFiliale = "FL-0008"
            
                Case Is = "Magdeburg"
                    retFiliale = "FL-0009"

                Case Is = "Mainz"
                    retFiliale = "FL-0010"
            
                Case Is = "München"
                    retFiliale = "FL-0011"
            
                Case Is = "Paderborn"
                    retFiliale = "FL-0012"
                    
                Case Is = "Regensburg"
                    retFiliale = "FL-0013"
            
                Case Is = "Rosenheim"
                    retFiliale = "FL-0014"
            
                Case Is = "Stuttgart"
                    retFiliale = "FL-0015"
             End Select
             
            id_Filiale = retFiliale

End Function

Function id_Tag(Tag As String) As String
        
        Dim retTag As String
        
        If Len(Tag) = 1 Then
            retTag = "TG-0" & Tag
        
        Else
            retTag = "TG-" & Tag
        End If
        
        id_Tag = retTag
End Function

Function id_Monat(Monat As String) As String
        
            Dim retMonat As String
            
            Select Case Monat
                Case Is = "1", "Januar"
                    retMonat = "M1-01"
                    
                Case Is = "2", "Februar"
                    retMonat = "M1-02"
                    
                Case Is = "3", "März"
                    retMonat = "M1-03"
                    
                Case Is = "4", "April"
                    retMonat = "M1-04"
                    
                Case Is = "5", "Mai"
                    retMonat = "M1-05"
                    
                Case Is = "6", "Juni"
                    retMonat = "M1-06"
                    
                Case Is = "7", "Juli"
                    retMonat = "M1-07"
                    
                Case Is = "8", "August"
                    retMonat = "M1-08"
                    
                Case Is = "9", "September"
                    retMonat = "M1-09"
                    
                Case Is = "10", "Oktober"
                    retMonat = "M1-10"
                    
                Case Is = "11", "November"
                    retMonat = "M1-11"
                    
                Case Is = "12", "Dezember"
                    retMonat = "M1-12"
                
            End Select
            
            id_Monat = retMonat
End Function

Function tageAufLager2(VKDate As Date, EKDate As Date) As Double
        
        Dim retTage As Double
        
        Dim EKTag As Double
        Dim EKMonat As Double
        Dim EKJahr As Double
        
        Dim VKTag As Double
        Dim VKMonat As Double
        Dim VKJahr As Double
        
        
                    'Tag Einkauf
                    EKTag = Val(Day(EKDate))
                    
                    'Tag Einkauf
                    EKMonat = Val(Month(EKDate))
                    
                    'Tag Einkauf
                    EKJahr = Val(Year(EKDate))
                    
                    'MsgBox VKDate
                     If Val(Year(VKDate)) = 9999 Then
                        
                        VKDate = Date
                        'MsgBox VKDate
                    End If
                   
                    'Tag Verkauf
                    VKTag = Val(Day(VKDate))
                    
                    'Tag Verkauf
                    VKMonat = Val(Month(VKDate))
                    
                    'Tag Verkauf
                    VKJahr = Val(Year(VKDate))
                    
                    
                    retTage = ((VKJahr * 365) - (EKJahr * 365)) + ((VKMonat * 12) - (EKMonat * 12)) + ((VKTag - EKTag))
        
        tageAufLager = retTage
End Function

Function tageAufLager(VKDate As Date, EKDate As Date) As Double
        
        Dim retTage As Double
        
        Dim EKTag As Double
        Dim EKMonat As Double
        Dim EKJahr As Double
        
        Dim VKTag As Double
        Dim VKMonat As Double
        Dim VKJahr As Double
        
        
                    
                    
        'MsgBox VKDate
            If Val(Year(VKDate)) = 9999 Then
                        
                VKDate = Date
                
            End If
            retTage = DateDiff("d", EKDate, VKDate)
                           
            tageAufLager = retTage
End Function

Function fkt_LagerbestandSeg()
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Lager As Worksheet
    Set Lager = Worksheets("Lager")
    
    Dim LagerOG As Worksheet
    Set LagerOG = Worksheets("Lager(OG)")
    
    Dim LagerBS As Worksheet
    Set LagerBS = Worksheets("Lagebestand")
    
    Dim a As Integer
    Dim b As Integer
    
    a = Lager.UsedRange.Rows.Count
    b = Lager.UsedRange.Columns.Count
    
    Dim Tag As String
    Dim Monat As String
    Dim Jahr As String
    
    Dim zwischenDate As Date
    zwischenDate = "01.07.2018"

    Dim land As String
    Dim marke As String
    
    Dim Summe1 As Integer
    Summe1 = 0
    
    Dim Summe2 As Integer
    Summe2 = 0
    
    Dim Summe3 As Integer
    Summe3 = 0
    
    'Für Jedes Auto in Lager(OG)
    For i = 2 To a
        
      If Lager.Cells(i, 12).Value < zwischenDate And zwischenDate < Lager.Cells(i, 13).Value Then
        If Lager.Cells(i, 8).Value = Worksheets("Preissegment ID").Cells(2, 1).Value Then
            Summe1 = Summe1 + 1
            
        End If
        
        If Lager.Cells(i, 8).Value = Worksheets("Preissegment ID").Cells(3, 1).Value Then
            Summe2 = Summe2 + 1
            
        End If
        
        If Lager.Cells(i, 8).Value = Worksheets("Preissegment ID").Cells(4, 1).Value Then
            Summe3 = Summe3 + 1

        End If
      End If

    Next i
    LagerBS.Cells(2, 4).Value = Summe1
    LagerBS.Cells(3, 4).Value = Summe2
    LagerBS.Cells(4, 4).Value = Summe3

End Function


Function fkt_ZeitAufLager()
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Lager As Worksheet
    Set Lager = Worksheets("Lager")
    
    Dim LagerOG As Worksheet
    Set LagerOG = Worksheets("Lager(OG)")
    
    Dim LagerZT As Worksheet
    Set LagerZT = Worksheets("Zeit auf Lager")
    
    Dim MarkeID As Worksheet
    Set MarkeID = Worksheets("Marke ID")
    
    Dim a As Integer
    Dim b As Integer
    
    a = Lager.UsedRange.Rows.Count
    b = Lager.UsedRange.Columns.Count
    
    Dim Tag As String
    Dim Monat As String
    Dim Jahr As String
    
    
    Dim Summe1 As Integer
    Summe1 = 0
    Dim TagesSumme1 As Integer
    TagesSumme1 = 0
    
    Dim Summe2 As Integer
    Summe2 = 0
    Dim TagesSumme2 As Integer
    TagesSumme2 = 0
    
    Dim Summe3 As Integer
    Summe3 = 0
    Dim TagesSumme3 As Integer
    TagesSumme3 = 0
    
    Dim Summe4 As Integer
    Summe4 = 0
    Dim TagesSumme4 As Integer
    TagesSumme4 = 0
    
    Dim Summe5 As Integer
    Summe5 = 0
    Dim TagesSumme5 As Integer
    TagesSumme5 = 0
    
    Dim Summe6 As Integer
    Summe6 = 0
    Dim TagesSumme6 As Integer
    TagesSumme6 = 0
    
    Dim Summe7 As Integer
    Summe7 = 0
    Dim TagesSumme7 As Integer
    TagesSumme7 = 0
    
    Dim Summe8 As Integer
    Summe8 = 0
    Dim TagesSumme8 As Integer
    TagesSumme8 = 0
    
    Dim Summe9 As Integer
    Summe9 = 0
    Dim TagesSumme9 As Integer
    TagesSumme9 = 0
    
    Dim Summe10 As Integer
    Summe10 = 0
    Dim TagesSumme10 As Integer
    TagesSumme10 = 0
    
    'Für Jedes Auto in Lager
    For i = 2 To a
        
      If Year(Lager.Cells(i, 12).Value) = 2018 And Year(Lager.Cells(i, 13).Value) = 2018 Then

        
            Select Case Lager.Cells(i, 2).Value
            
                Case Is = MarkeID.Cells(2, 1).Value
                    Summe1 = Summe1 + 1
                    TagesSumme1 = TagesSumme1 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(3, 1).Value
                    Summe2 = Summe2 + 1
                    TagesSumme2 = TagesSumme2 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(4, 1).Value
                    Summe3 = Summe3 + 1
                    TagesSumme3 = TagesSumme3 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(5, 1).Value
                    Summe4 = Summe4 + 1
                    TagesSumme4 = TagesSumme4 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(6, 1).Value
                    Summe5 = Summe5 + 1
                    TagesSumme5 = TagesSumme5 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(7, 1).Value
                    Summe6 = Summe6 + 1
                    TagesSumme6 = TagesSumme6 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(8, 1).Value
                    Summe7 = Summe7 + 1
                    TagesSumme7 = TagesSumme7 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(9, 1).Value
                    Summe8 = Summe8 + 1
                    TagesSumme8 = TagesSumme8 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(10, 1).Value
                    Summe9 = Summe9 + 1
                    TagesSumme9 = TagesSumme9 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(11, 1).Value
                    Summe10 = Summe10 + 1
                    TagesSumme10 = TagesSumme10 + Lager.Cells(i, 21).Value

            End Select
        
        
        
        
      End If

    Next i
    
    If Summe1 <> 0 Then
        LagerZT.Cells(2, 3).Value = TagesSumme1 / Summe1
    End If
    
    If Summe2 <> 0 Then
        LagerZT.Cells(3, 3).Value = TagesSumme2 / Summe2
    End If
    
    If Summe3 <> 0 Then
        LagerZT.Cells(4, 3).Value = TagesSumme3 / Summe3
    End If
    
    If Summe4 <> 0 Then
        LagerZT.Cells(5, 3).Value = TagesSumme4 / Summe4
    End If
    
    If Summe5 <> 0 Then
        LagerZT.Cells(6, 3).Value = TagesSumme5 / Summe5
    End If
    
    If Summe6 <> 0 Then
        LagerZT.Cells(7, 3).Value = TagesSumme6 / Summe6
    End If
    
    If Summe7 <> 0 Then
        LagerZT.Cells(8, 3).Value = TagesSumme7 / Summe7
    End If
    
    If Summe8 <> 0 Then
        LagerZT.Cells(9, 3).Value = TagesSumme8 / Summe8
    End If
    
    If Summe9 <> 0 Then
        LagerZT.Cells(10, 3).Value = TagesSumme9 / Summe9
    End If
    
    If Summe10 <> 0 Then
        LagerZT.Cells(11, 3).Value = TagesSumme10 / Summe10
    End If

End Function


Function fkt_AutoAlter()
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Lager As Worksheet
    Set Lager = Worksheets("Lager")
    
    Dim LagerOG As Worksheet
    Set LagerOG = Worksheets("Lager(OG)")
    
    Dim LagerZT As Worksheet
    Set LagerZT = Worksheets("Zeit auf Lager")
    
    Dim MarkeID As Worksheet
    Set MarkeID = Worksheets("Marke ID")
    
    Dim AutoA As Worksheet
    Set AutoA = Worksheets("Auto-Alter")
    
    Dim a As Integer
    Dim b As Integer
    
    a = Lager.UsedRange.Rows.Count
    b = Lager.UsedRange.Columns.Count
    
    Dim Tag As String
    Dim Monat As String
    Dim Jahr As String
    
    
    Dim Summe1 As Integer
    Summe1 = 0
    Dim TagesSumme1 As Integer
    TagesSumme1 = 0
    
    Dim Summe2 As Integer
    Summe2 = 0
    Dim TagesSumme2 As Integer
    TagesSumme2 = 0
    
    Dim Summe3 As Integer
    Summe3 = 0
    Dim TagesSumme3 As Integer
    TagesSumme3 = 0
    
    Dim Summe4 As Integer
    Summe4 = 0
    Dim TagesSumme4 As Integer
    TagesSumme4 = 0
    
    Dim Summe5 As Integer
    Summe5 = 0
    Dim TagesSumme5 As Integer
    TagesSumme5 = 0
    
    Dim Summe6 As Integer
    Summe6 = 0
    Dim TagesSumme6 As Integer
    TagesSumme6 = 0
    
    Dim Summe7 As Integer
    Summe7 = 0
    Dim TagesSumme7 As Integer
    TagesSumme7 = 0
    
    Dim Summe8 As Integer
    Summe8 = 0
    Dim TagesSumme8 As Integer
    TagesSumme8 = 0
    
    Dim Summe9 As Integer
    Summe9 = 0
    Dim TagesSumme9 As Integer
    TagesSumme9 = 0
    
    Dim Summe10 As Integer
    Summe10 = 0
    Dim TagesSumme10 As Integer
    TagesSumme10 = 0
    
    'Für Jedes Auto in Lager
    For i = 2 To a
        
      'Verkauft != Jahr 9999
      If Year(Lager.Cells(i, 13).Value) <> 9999 Then
        

            Select Case Lager.Cells(i, 2).Value
            
                Case Is = MarkeID.Cells(2, 1).Value
                    Summe1 = Summe1 + 1
                    TagesSumme1 = TagesSumme1 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(3, 1).Value
                    Summe2 = Summe2 + 1
                    TagesSumme2 = TagesSumme2 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(4, 1).Value
                    Summe3 = Summe3 + 1
                    TagesSumme3 = TagesSumme3 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(5, 1).Value
                    Summe4 = Summe4 + 1
                    TagesSumme4 = TagesSumme4 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(6, 1).Value
                    Summe5 = Summe5 + 1
                    TagesSumme5 = TagesSumme5 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(7, 1).Value
                    Summe6 = Summe6 + 1
                    TagesSumme6 = TagesSumme6 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(8, 1).Value
                    Summe7 = Summe7 + 1
                    TagesSumme7 = TagesSumme7 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(9, 1).Value
                    Summe8 = Summe8 + 1
                    TagesSumme8 = TagesSumme8 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(10, 1).Value
                    Summe9 = Summe9 + 1
                    TagesSumme9 = TagesSumme9 + Lager.Cells(i, 21).Value
                    
                Case Is = MarkeID.Cells(11, 1).Value
                    Summe10 = Summe10 + 1
                    TagesSumme10 = TagesSumme10 + Lager.Cells(i, 21).Value

            End Select
        
        
        
        
      End If

    Next i
    
    If Summe1 <> 0 Then
        AutoA.Cells(2, 3).Value = TagesSumme1 / Summe1
    End If
    
    If Summe2 <> 0 Then
        AutoA.Cells(3, 3).Value = TagesSumme2 / Summe2
    End If
    
    If Summe3 <> 0 Then
        AutoA.Cells(4, 3).Value = TagesSumme3 / Summe3
    End If
    
    If Summe4 <> 0 Then
        AutoA.Cells(5, 3).Value = TagesSumme4 / Summe4
    End If
    
    If Summe5 <> 0 Then
        AutoA.Cells(6, 3).Value = TagesSumme5 / Summe5
    End If
    
    If Summe6 <> 0 Then
        AutoA.Cells(7, 3).Value = TagesSumme6 / Summe6
    End If
    
    If Summe7 <> 0 Then
        AutoA.Cells(8, 3).Value = TagesSumme7 / Summe7
    End If
    
    If Summe8 <> 0 Then
        AutoA.Cells(9, 3).Value = TagesSumme8 / Summe8
    End If
    
    If Summe9 <> 0 Then
        AutoA.Cells(10, 3).Value = TagesSumme9 / Summe9
    End If
    
    If Summe10 <> 0 Then
        AutoA.Cells(11, 3).Value = TagesSumme10 / Summe10
    End If

End Function


