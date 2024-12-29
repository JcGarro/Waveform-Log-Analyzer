Attribute VB_Name = "Module3"
Public ThisW As Worksheet
Public Sub Proce()
   Dim test1 As Variant
   Dim test2 As Variant
   Dim ComS As Worksheet
   Dim MType As String
   Dim ResAdd As String
   Dim CurrStat As String
   Dim E As Integer
   Dim A As Integer
   Dim Rr As Integer
   Dim cnt As Integer
   Dim adde As Integer
   Dim iRange As Range
   Dim iCells As Range
   Dim events As Variant
   
   '--------------------------------Logs
   Dim Va As Variant
   Dim Vb As Variant
   Dim Vc As Variant
   Dim Ia As Variant
   Dim Ib As Variant
   Dim Ic As Variant
   Dim Bas As Range
   Dim PQs As Range
   
   '--------------------------------Error
   Dim timeE As Integer
   Dim cdurE As Integer
   Dim wdurE As Integer
   
   '--------------------------------Coordinates
   Dim Cw As Integer
   Dim Rw As Integer
   Dim Cp As Integer
   Dim Rp As Integer
   
   Dim Val As Integer
   Dim Addf As Integer
   Dim Addg As Integer
   Dim Addh As Integer
   Dim Start As Integer


   Dim ebin As Integer
   Dim even As Integer
   Dim durz As String
   
   If Intersect(ActiveCell, Range("C20:C35")) Is Nothing Then
        MsgBox "Select the Meter Name to proceed. Try Again."
        Exit Sub
   End If
   
   cnt = Application.WorksheetFunction.CountIf(ActiveWorkbook.Sheets("Main").Range("K8:K35"), "*=*") 'events count
   CurrStat = ActiveCell.Offset(, 2).Value
   
   If CurrStat = "Stored" Then
        MType = ActiveCell.Offset(0, 1).Value
        ResAdd = ActiveCell.Offset(, 2).Address()
   Else
        MsgBox "No data to be processed."
        Exit Sub
   End If
   
   Set ComS = ActiveWorkbook.Sheets("Comparison Table")
   Set ThisW = ActiveWorkbook.Sheets(ActiveCell.Value)
   Set Ma = ActiveWorkbook.Sheets("Main")
   
   Call Setmeters
   
   '---------------------------Number of row and cols
   Cw = 17
   Rw = ThisW.Range("K1").Value
   Cp = 7
   Rp = ThisW.Range("N1").Value
   
   
   '----------------------------------------------------------------------Wave data
   Set Bas = ThisW.Range(Cells(3, 10).Address(), Cells(Rw + 1, 10).Address())
   events = Bas.Offset(, 4)
   Tri = Bas.Offset(, 2)
   Durr = Bas.Offset(, 3)
   Va = Bas.Offset(, 6)
   Vb = Bas.Offset(, 7)
   Vc = Bas.Offset(, 8)
   Ia = Bas.Offset(, 13)
   Ib = Bas.Offset(, 14)
   Ic = Bas.Offset(, 15)
   
   '----------------------------------------------------------------------PQ data
   Set PQs = ThisW.Range(Cells(Rw + 4, 10).Address(), Cells(Rw + Rp + 2, 10).Address())
   StTime = PQs
   Durp = PQs.Offset(, 2)
   
   
   Rr = 0
   
   For even = 1 To UBound(events)
        
        If Right(events(even, 1), 6) <> "Normal" And events(even, 1) <> "" Then
             Rr = Rr + 1
        End If
        
   Next even
  '-----------------------------------------------------------------------------------------------------------------------------------------------------Pasting Table
      
   '-------------------------------------------------------------------------------------------------------------------- Input Events
   
   ThisW.Range(CoAd(0, 0), CoAd(0, 1)).Interior.Color = RGB(256, 236, 156)
   ThisW.Range(CoAd(0, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(0, 0)).Value = "Input"
   ThisW.Range(CoAd(0, 1)).Value = "Event | Amp % (ms)"
   
   
   
   
   adde = 1
   
   For evin = 1 To cnt
        Dim percen As String
        
        If Mid(Ma.Range(Cells(7 + evin, 11).Address()).Value, 4, 3) <> "Nor" Then
            durz = " (" & CStr(Ma.Range("D10").Value) & ")"
        Else
            durz = ""
        End If
        
        percen = Application.WorksheetFunction.Text(Ma.Range(Cells(7 + evin, 9).Address()).Value, "0.0%")
        
        
        If DUT(MType).CurrentSag = True Then
            If Mid(Ma.Range(Cells(7 + evin, 11).Address()).Value, 4, 3) <> "Nor" Then
                ThisW.Range(CoAd(0, 1 + evin)).Value = Ma.Range(Cells(7 + evin, 11).Address()).Value & " | " & percen & durz
            ElseIf True Then
                ThisW.Range(CoAd(0, 1 + evin)).Value = "Normal"
            End If
        
        Else
            If Mid(Ma.Range(Cells(7 + evin, 11).Address()).Value, 4, 3) = "Sur" Then
                ThisW.Range(CoAd(0, 1 + adde)).Value = Ma.Range(Cells(7 + evin, 11).Address()).Value & " | " & percen & durz
                adde = adde + 1
            ElseIf Mid(Ma.Range(Cells(7 + evin, 11).Address()).Value, 4, 3) = "Sag" Then
                If Left(Ma.Range(Cells(7 + evin, 11).Address()).Value, 1) <> "I" Then
                    ThisW.Range(CoAd(0, 1 + adde)).Value = Ma.Range(Cells(7 + evin, 11).Address()).Value & " | " & percen & durz
                    adde = adde + 1
                End If
            End If
            
            'MsgBox adde
        End If
   Next evin
   
   '--------------------------------- Boarders
   
   If DUT(MType).CurrentSag = False Then
        cnt = adde - 1
   End If
    
   Set iRange = ThisW.Range(CoAd(0, 0), CoAd(6, cnt + 1))

   For Each iCells In iRange
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
   Next iCells

   ThisW.Columns("A:I").HorizontalAlignment = xlCenter
   ThisW.Columns("D").NumberFormat = "m/d/yyyy  h:mm:ss.000"
   ThisW.Columns("F").NumberFormat = "m/d/yyyy  h:mm:ss.000"

    
    '------------------------------------------------------------------------------------------------------------------Limits
   ThisW.Range(CoAd(1, 0), CoAd(1, 1)).Interior.Color = RGB(256, 236, 156)
   ThisW.Range(CoAd(1, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(1, 0)).Value = "Wave Boundary"
   ThisW.Range(CoAd(1, 1)).Value = "Upper | Lower"
   ThisW.Range(CoAd(1, 3)).Font.Color = vbBlack
   
   
   For ebin = 1 To cnt
        If Left(ThisW.Range(CoAd(0, ebin + 1)), 1) = "V" And Mid(ThisW.Range(CoAd(0, ebin + 2)), 4, 3) <> "Nor" Then
            ThisW.Range(CoAd(1, ebin + 1)).Value = DUT(MType).VSurge & "% | " & DUT(MType).VSag & "%"
        ElseIf Left(ThisW.Range(CoAd(0, ebin + 1)), 1) = "I" And Mid(ThisW.Range(CoAd(0, ebin + 2)), 4, 3) <> "Nor" Then
            ThisW.Range(CoAd(1, ebin + 1)).Value = DUT(MType).ISurge & "% | " & DUT(MType).ISag & "%"
        End If
   Next ebin

   '-------------------------Getting 1st non normal event or start of actual events
   Addf = 0
   Start = 1

   For Evt = 1 To UBound(events)
        If InStr(events(Evt, 1), "Normal") <> 0 Then
            Addf = Addf + 1
            
        Else
            Exit For
        End If
   Next Evt
   
   If DUT(MType).Normal = True Then
        Start = Start + Addf
   End If


  '---------------------------------------------------------------------------------------------------------------Trigger Time
   ThisW.Range(CoAd(2, 0), CoAd(2, 1)).Interior.Color = RGB(208, 236, 252)
   ThisW.Range(CoAd(2, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(2, 0)).Value = "Trigger Time (W)"
   ThisW.Range(CoAd(2, 1)).Value = "Date & Time"
   ThisW.Range(CoAd(2, 3)).Font.Color = vbBlack

   
   If Normals(MType) = False Then
        For tt = 1 To UBound(Tri)
            If Tri(tt, 1) <> "" And tt <= cnt Then
                ThisW.Range(CoAd(2, 1 + tt)).Value = Tri(tt, 1)
            End If
        Next tt
   Else
        'Laters
        Dim Vs As Integer
        Vs = 1
        For tta = Start To UBound(Tri)
            If Tri(tta, 1) <> "" And tta <= cnt Then
                ThisW.Range(CoAd(2, 1 + Vs)).Value = Tri(tta, 1)
                Vs = Vs + 1
            End If
        Next tta
        
   End If
   
   '--------------------------------------------------------------------------------------------------------------- Actual
   Dim eve As Integer
   ThisW.Range(CoAd(3, 0), CoAd(3, 1)).Interior.Color = RGB(208, 236, 252)
   ThisW.Range(CoAd(3, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(3, 0)).Value = "Waveforms (W)"
   ThisW.Range(CoAd(3, 1)).Value = "Events | Amp %"
   ThisW.Range(CoAd(3, 3)).Font.Color = vbBlack
      
  '----------------------------------------------------------------
   Val = 0
   For eve = Start To UBound(events)
        
        If events(eve, 1) <> "" And Val < cnt Then
             
             Val = Val + 1
             Dim Valu As String
             If Left(events(eve, 1), 1) = "V" Then
                If Left(events(eve, 1), 2) = "Va" Then
                    Valu = Application.WorksheetFunction.Text((Va(eve, 1) / 120), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                ElseIf Left(events(eve, 1), 2) = "Vb" Then
                    Valu = Application.WorksheetFunction.Text((Vb(eve, 1) / 120), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                ElseIf Left(events(eve, 1), 2) = "Vc" Then
                    Valu = Application.WorksheetFunction.Text((Vc(eve, 1) / 120), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                End If
             ElseIf Left(events(eve, 1), 1) = "I" Then
                If Left(events(eve, 1), 2) = "Ia" Then
                    Valu = Application.WorksheetFunction.Text((Ia(eve, 1) / 5), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                ElseIf Left(events(eve, 1), 2) = "Ib" Then
                    Valu = Application.WorksheetFunction.Text((Ib(eve, 1) / 5), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                ElseIf Left(events(eve, 1), 2) = "Ic" Then
                    Valu = Application.WorksheetFunction.Text((Ic(eve, 1) / 5), "0.0%")
                    ThisW.Range(CoAd(3, 1 + Val)).Value = events(eve, 1) & " | " & Valu
                End If
             Else
                'ThisW.Range(Cells(2 + Val, 3).Address()).Value = "I"
             End If
             
        End If
        
   Next eve
  
  
    
  
  

  '--------------------------------------------------------------------------------------------------------------Capture Begin
   ThisW.Range(CoAd(4, 0), CoAd(4, 1)).Interior.Color = RGB(208, 236, 252)
   ThisW.Range(CoAd(4, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(4, 0)).Value = "Capture Begin (PQ)"
   ThisW.Range(CoAd(4, 1)).Value = "Date & Time"
   ThisW.Range(CoAd(4, 3)).Font.Color = vbBlack
   
   If Normals(MType) = False Then
        Dim Adda As Integer
        For cb = 1 To UBound(StTime)
            If StTime(cb, 1) <> "" And Durp(cb, 1) <> 0 Then '-----------------Only PQ with Dur will enter
                Adda = 1
                For aa = 1 To UBound(Tri) '-------------------Finding Address of Match
                    If Tri(aa, 1) = StTime(cb, 1) And Adda <= cnt Then
                        ThisW.Range(CoAd(4, Adda + 1)).Value = StTime(cb, 1)
                    End If
                    Adda = Adda + 1
                Next aa
            End If
        Next cb
   Else
        For cb = 1 To UBound(StTime)
'            If StTime(cb, 1) <> "" And Durp(cb, 1) <> 0 Then '-----------------Only PQ with Dur will enter
'                Addg = 0
'                For aa = Start To UBound(Tri) '-------------------Finding Address of Match
'                    If Tri(aa, 1) = StTime(cb, 1) And Addg <= cnt Then
'                        ThisW.Range(CoAd(3, Addg + 2)).Value = StTime(cb, 1)
'                    End If
'                    Addg = Addg + 1
'                Next aa
'            End If
            If StTime(cb, 1) <> "" Then '-----------------Only PQ with Dur will enter
                        ThisW.Range(CoAd(4, 2 * cb)).Value = StTime(cb, 1)
            End If
        
        Next cb
   
   End If
  
  '----------------------------------Capture Duration
   ThisW.Range(CoAd(5, 0), CoAd(5, 1)).Interior.Color = RGB(208, 236, 252)
   ThisW.Range(CoAd(5, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(5, 0)).Value = "Capture Duration (W) "
   ThisW.Range(CoAd(5, 1)).Value = "(ms)"
   ThisW.Range(CoAd(5, 3)).Font.Color = vbBlack
  
    Addh = 1
        For cd = Start To UBound(Durr)
            If Durr(cd, 1) <> "" And cd <= cnt Then
                ThisW.Range(CoAd(5, 1 + Addh)).Value = Durr(Addh, 1)
                Addh = Addh + 1
            End If
        Next cd
  
  '----------------------------------Wave Duration
   ThisW.Range(CoAd(6, 0), CoAd(6, 1)).Interior.Color = RGB(208, 236, 252)
   ThisW.Range(CoAd(6, 2)).Interior.Color = RGB(255, 255, 255)
   ThisW.Range(CoAd(6, 0)).Value = "Wave Duration (PQ)"
   ThisW.Range(CoAd(6, 1)).Value = "(ms)"
   ThisW.Range(CoAd(6, 3)).Font.Color = vbBlack
  
   If Normals(MType) = False Then
        Dim Addb As Integer
        For wd = 1 To UBound(Durp)
            If Durp(wd, 1) <> "" And Durp(wd, 1) <> 0 Then '---------------Only PQ with Non zero Dur will enter
                Addb = 1
                For ab = 1 To UBound(Tri) '-------------------Finding Address of Match
                    If Tri(ab, 1) = StTime(wd, 1) And Addb <= cnt Then
                        ThisW.Range(CoAd(6, Addb + 1)).Value = Durp(wd, 1)
                    End If
                    Addb = Addb + 1
                Next ab
                'ThisW.Range(CoAd(5, 1 + wd)).Value = Durp(wd, 1)
            End If
        Next wd
   Else
        For wd = 1 To UBound(Durp)
                    If Durp(wd, 1) <> "" Then
                        ThisW.Range(CoAd(6, wd * 2)).Value = Durp(wd, 1)
                    End If
        Next wd

   End If
  
  '---------------------------------------------------------------------------------------------------------------------------Evaluating Part
  
  '------------------------Actual Events
  Call Colorize(MType, cnt)

  
  '----------------------- Trigger Time vs PQ Begin
  For ae = 1 To cnt
    If ThisW.Range(CoAd(2, 1 + ae)).Value = ThisW.Range(CoAd(4, 1 + ae)).Value And ThisW.Range(CoAd(4, 1 + ae)).Value <> "" Then
        ThisW.Range(CoAd(4, 1 + ae)).Interior.Color = RGB(144, 238, 144)
    Else
        If ThisW.Range(CoAd(4, 1 + ae)).Value <> "" Then
            timeE = timeE + 1
            ThisW.Range(CoAd(4, 1 + ae)).Interior.Color = RGB(205, 92, 92)
        End If
    End If
  Next ae
  '----------------------- Capture Durations
  For Ac = 1 To cnt
      If ThisW.Range(CoAd(5, 1 + Ac)).Value = DUT(MType).Duration Then
        ThisW.Range(CoAd(5, 1 + Ac)).Interior.Color = RGB(144, 238, 144)
      Else
        'Count the error
        If ThisW.Range(CoAd(5, 1 + Ac)).Value <> 0 Then
            cdurE = cdurE + 1
            ThisW.Range(CoAd(5, 1 + Ac)).Interior.Color = RGB(205, 92, 92)
        End If
      End If
  Next Ac
  
  '----------------------- Wave Durations
  Dim TThres
  TThres = 10
  For ad = 1 To cnt
    If Abs(ThisW.Range(CoAd(6, 1 + ad)).Value - Ma.Range("D10").Value) < TThres Then
        ThisW.Range(CoAd(6, 1 + ad)).Interior.Color = RGB(144, 238, 144)
    Else
        'Count the error
        If ThisW.Range(CoAd(6, 1 + ad)).Value <> "" Then
            wdurE = wdurE + 1
            ThisW.Range(CoAd(6, 1 + ad)).Interior.Color = RGB(205, 92, 92)
        End If
    End If
  Next ad
  
  ThisW.Range(CoAd(0, 0), CoAd(6, 1)).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
  ThisW.Range(CoAd(0, 0), CoAd(6, cnt + 1)).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
  ThisW.Columns("A:I").AutoFit
  ThisW.Range(Cells(3, 10).Address(), Cells(Rw + 1, 12).Address()).NumberFormat = "m/d/yyyy h:mm:ss.000"
  ThisW.Range(Cells(Rw + 4, 10).Address(), Cells(Rw + Rp + 2, 11).Address()).NumberFormat = "m/d/yyyy h:mm:ss.000"
  '---------------------------------------------------------------------------------------------------Decisioning
  If eventE = 0 And timeE = 0 And cdurE = 0 And wdurE = 0 Then ' Satifactory
    Ma.Range(ResAdd).Value = "Satisfactory"
  Else
    Ma.Range(ResAdd).Value = "With Issues"
    Ma.Range(ResAdd).Offset(, 1).Value = CStr(eventE) & "     |     " & CStr(timeE) & "    |     " & CStr(cdurE) & "       |     " & CStr(wdurE)
  End If
   

End Sub
