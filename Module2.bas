Attribute VB_Name = "Module2"
Public N1500P As Meters
Public N1450 As Meters
Public S200 As Meters
Public S250 As Meters
Public S270 As Meters

Public Sub Setmeters()
        Set N1500P = New Meters
        With N1500P
            '.Normal = False 'Muna
            .Normal = True
            .CurrentSag = True
            .Duration = 3000
            .VSurge = 110
            .VSag = 90
            .ISurge = 200
            .ISag = 0
            
        End With
        Set S200 = New Meters
        With S200
            .Normal = False
            .CurrentSag = False
            .Duration = 533
            .VSurge = 120
            .VSag = 80
            .ISurge = 200
            .ISag = 0
            
        End With
        Set S250 = New Meters
        With S250
            .Normal = False
            .CurrentSag = False
            .Duration = 533
            .VSurge = 120
            .VSag = 80
            .ISurge = 200
            .ISag = 0
            
        End With
        
        Set S270 = New Meters
        With S270
            .Normal = False
            .CurrentSag = False
            .Duration = 533
            .VSurge = 120
            .VSag = 80
            .ISurge = 200
            .ISag = 0
            
        End With
        
        
End Sub

Public Sub Colorize(MType As String, cnt As Integer)
  Dim Unit As Meters
  eventE = 0
  If MType = "Nexus 1500+" Then
    Set Unit = N1500P
  ElseIf MType = "Nexus 1450" Then
    Set Unit = N1450
  ElseIf MType = "Shark 200" Then
    Set Unit = S200
  ElseIf MType = "Shark 250" Then
    Set Unit = S250
  ElseIf MType = "Shark 270" Then
    Set Unit = S700
  End If
  
  
  For i = 1 To cnt
        Dim Act As Double
        Dim Exp As Double
        Dim StatA As String
        Dim StatE As String
        Dim StatAf As String
        Dim StatEf As String
        Dim Param
        Dim Thres As Double
        
        Thres = CDbl(ActiveWorkbook.Sheets("Main").Range("D11").Value) * 100
        
        If ThisW.Range(Cells(2 + i, 2).Address()).Value <> "" Then
            If Left(ThisW.Range(Cells(2 + i, 2).Address()).Value, 1) = "V" Then
                If Mid(ThisW.Range(Cells(2 + i, 2).Address()).Value, 4, 3) = "Sur" Then
                    Exp = DUT(MType).VSurge
                ElseIf Mid(ThisW.Range(Cells(2 + i, 2).Address()).Value, 4, 3) = "Sag" Then
                    Exp = DUT(MType).VSag
                End If
            ElseIf Left(ThisW.Range(Cells(2 + i, 2).Address()).Value, 1) = "I" Then
                If Mid(ThisW.Range(Cells(2 + i, 2).Address()).Value, 4, 3) = "Sur" Then
                    Exp = DUT(MType).ISurge
                ElseIf Mid(ThisW.Range(Cells(2 + i, 2).Address()).Value, 4, 3) = "Sag" Then
                    Exp = DUT(MType).ISag
                End If
            End If
            StatEf = Split(ThisW.Range(Cells(2 + i, 2).Address()).Value, "|")(0)
            
        Else
            StatE = ""
        End If
        
        If ThisW.Range(Cells(2 + i, 5).Address()).Value <> "" Then
            Act = CDbl(Replace(Split(ThisW.Range(Cells(2 + i, 5).Address()).Value, "|")(1), "%", "")) 'Value%
            Param = Left(ThisW.Range(Cells(2 + i, 5).Address()).Value, 1) 'Volts or Cur
            StatA = Mid(ThisW.Range(Cells(2 + i, 5).Address()).Value, 4, 3) 'Sag or Surge
            StatAf = Split(ThisW.Range(Cells(2 + i, 5).Address()).Value, "|")(0) 'All
        Else
            StatAf = ""
        End If
        
        
        If StatEf <> StatAf Then
            If StatAf <> "" Then
                If StatA = "Nor" Then
                    If (100 - Act) < Thres Or (Act - 100) < Thres Then
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(144, 238, 144)
                    End If
                Else
                    ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(205, 92, 92)
                    eventE = eventE + 1
                End If
            End If
        ElseIf StatEf = StatAf Then
            If StatA = "Sur" Then
                If Param = "V" Then
                    If (Act - Unit.VSurge) < Thres Then
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(144, 238, 144) 'Green
                    Else
                       ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(205, 92, 92) 'Red
                       eventE = eventE + 1
                    End If
                ElseIf Param = "I" Then
                    If (Act - Unit.ISurge) < Thres Then
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(144, 238, 144)
                    Else
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(205, 92, 92)
                        eventE = eventE + 1
                    End If
                End If
            ElseIf StatA = "Sag" Then
                If Param = "V" Then
                    If (Unit.VSag - Act) < Thres Then
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(144, 238, 144)
                    Else
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(205, 92, 92)
                        eventE = eventE + 1
                    End If
                ElseIf Param = "I" Then
                    If (Unit.ISag - Act) < Thres Then
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(144, 238, 144)
                    Else
                        ThisW.Range(Cells(2 + i, 5).Address()).Interior.Color = RGB(205, 92, 92)
                        eventE = eventE + 1
                    End If
                End If
            End If
        End If
        
  Next i
End Sub

Public Function CoAd(Col As Integer, Row As Integer) As String
    CoAd = Range("B1").Offset(Row, Col).Address()
End Function
Public Function DUT(MType As String) As Meters
    Dim Unit As Meters
    If MType = "Nexus 1500+" Then
      Set Unit = N1500P
    ElseIf MType = "Nexus 1450" Then
      Set Unit = N1450
    ElseIf MType = "Shark 200" Then
      Set Unit = S200
    ElseIf MType = "Shark 250" Then
      Set Unit = S250
    ElseIf MType = "Shark 270" Then
      Set Unit = S270
    End If

    Set DUT = Unit
End Function
Public Function Normals(MType As String) As Boolean
    Dim Unit As Meters
    If MType = "Nexus 1500+" Then
      Set Unit = N1500P
    ElseIf MType = "Nexus 1450" Then
      Set Unit = N1450
    ElseIf MType = "Shark 200" Then
      Set Unit = S200
    ElseIf MType = "Shark 250" Then
      Set Unit = S250
    ElseIf MType = "Shark 270" Then
      Set Unit = S270
    End If

    Normals = Unit.Normal
End Function
