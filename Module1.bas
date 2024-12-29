Attribute VB_Name = "Module1"

Public Compa As Worksheet
Public Cont As Boolean 'Checker if there is an item in clipboard
Public PQq As Boolean
Public Cleardata As Boolean
Public MeterType As String
Public MeterName As String
Public StatAdd As String
Public eventE As Integer

'--------------------------Coords
Public Rs As Integer
Public Cs As Integer
Public Rz As Integer
Public Cz As Integer



Private Sub MeterCheck()
    
   MeterType = "" 'Making it empty first
  
  MeterType = ActiveWorkbook.Sheets("Main").Range(ActiveCell.Offset(, 1).Address).Value
   
End Sub
Public Sub ImpoW()

    '------------------------------------------- Initialization
    Dim CsvF As String 'csv dir
    Dim R0 As Integer
    Dim C0 As Integer
    Dim csvWB As Object 'csv obj
    Dim exApp As Object
    
    Set exApp = CreateObject("Excel.Application." & CLng(Application.Version))
                
    '-------------------------------------------Log Verification
    If ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Wave Only" Then 'PQ with Wave
        Call ImpoP
        PQq = True
        Exit Sub
    End If
    
    '------------------------------------------- Opening File
    CsvF = Application.GetOpenFilename()
    If CsvF = "False" Then
        MsgBox "No Selected File"
        Cont = False
        Exit Sub
    End If
    If (Right(CsvF, 4)) <> ".csv" And (Right(CsvF, 4)) <> "xlsx" Then
        MsgBox "Invalid File"
        Cont = False
        Exit Sub
    End If
    Set csvWB = exApp.Workbooks.Open(CsvF)
    
    csvWB.ActiveSheet.Activate
    csvWB.ActiveSheet.Range("A1").Activate
    
    '----------------------------------------- Copying File

    R0 = csvWB.ActiveSheet.UsedRange.Rows.Count
    C0 = csvWB.ActiveSheet.UsedRange.Columns.Count
    
    If C0 = 17 And ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "" Then 'Waveform
        csvWB.ActiveSheet.Range("A1", Cells(R0, C0).Address()).Copy
        Rs = R0
        Cs = C0
    ElseIf C0 = 7 And ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "" Then 'PQ without Wave
        MsgBox "Please import Waveform Logs first."
        Cont = False
        csvWB.Close
        exApp.Quit
        ActiveWorkbook.Sheets("Main").Activate
        Exit Sub
    Else
        MsgBox "Invalid log type."
        Cont = False
    End If

    csvWB.Close
    exApp.Quit
    Cont = True
    
   
End Sub
Public Sub ImpoP()

    '------------------------------------------- Initialization
    Dim CsvF1 As String 'csv dir
    Dim R1 As Integer
    Dim C1 As Integer
    Dim csvWB1 As Object 'csv obj
    Dim exApp1 As Object
    
    Set exApp1 = CreateObject("Excel.Application." & CLng(Application.Version))
    
    '------------------------------------------- Opening File
    CsvF1 = Application.GetOpenFilename()
    If CsvF1 = "False" Then
        MsgBox "No Selected File"
        Cont = False
        Exit Sub
    End If
    If (Right(CsvF1, 4)) <> ".csv" And (Right(CsvF1, 4)) <> "xlsx" Then
        MsgBox "Invalid File"
        Cont = False
        Exit Sub
    End If
    Set csvWB1 = exApp1.Workbooks.Open(CsvF1)
    
    csvWB1.ActiveSheet.Activate
    csvWB1.ActiveSheet.Range("A1").Activate
    Call MeterCheck
   
    '----------------------------------------- Copying File

    R1 = csvWB1.ActiveSheet.UsedRange.Rows.Count
    C1 = csvWB1.ActiveSheet.UsedRange.Columns.Count


    If C1 = 7 And ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Wave Only" Then  'PQ with Wave already
        csvWB1.ActiveSheet.Range("A1", Cells(R1, C1).Address()).Copy
        Rz = R1
        Cz = C1
    ElseIf C1 = 17 And ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Wave Only" Then 'Wave again
        MsgBox "This is another waveform logs."
        Cont = False
        Exit Sub
    Else
        MsgBox "This is not PQ Log."
        Cont = False
    End If
    
    csvWB1.Close
    exApp1.Quit
    Cont = True
    
   
End Sub

Public Sub WPaster()
    
    
    Set Compa = ActiveWorkbook.Sheets(MeterName)
    Dim sRange As Range
    
    '------------------- Pasting
    Compa.Activate
    Compa.Range("J2").Activate
    Compa.Paste
    Compa.Range("J2", Cells(Rs + 1, 27).Address()).Sort Key1:=Range("J2"), Order1:=xlAscending, Header:=xlYes
    Compa.Range("J1").Value = "Wave Log#:"
    'Compa.Range("K1").Value = Cs
    Compa.Range("K1").Value = Rs
    Compa.Range("P1").Value = "Meter Type"
    Compa.Range("Q1").Value = MeterType
    
    '------------------- Bordering
    
    Set sRange = Compa.Range("J2", Cells(Rs + 1, 26).Address())

    For Each iCells In sRange
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    Next iCells
    Compa.Range("J2", Cells(Rs + 1, 26).Address()).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
    Compa.Range("J2", Cells(2, 26).Address()).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
    
    
    '--------------------------------------Changing Status in Main sheet
    ActiveWorkbook.Sheets("Main").Activate
    ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Wave Only"
    
End Sub
Public Sub PPaster()
    Dim Pp
    Dim Rx As Integer
    Dim tRange As Range
    
    
    Set Pp = ActiveWorkbook.Sheets(MeterName)
    'Cx = Pp.Range("K1").Value
    Rx = Pp.Range("K1").Value
    
    
    '------------------- Pasting
    If Rx <> 0 Then
        Pp.Activate
        Pp.Range(Cells(Rx + 3, 10).Address()).Activate
        Pp.Paste
        Pp.Range(Cells(Rx + 3, 10).Address(), Cells(Rx + Rz + 2, 10 + Cz).Address()).Sort Key1:=Range(Cells(Rx + 3, 10).Address()), Order1:=xlAscending, Header:=xlYes
        Pp.Range("M1").Value = "PQ Log #:"
        'Pp.Range("N1").Value = Cz
        Pp.Range("N1").Value = Rz
    End If
    
    '------------------- Bordering
    
    Set tRange = Pp.Range(Cells(Rx + 3, 10).Address(), Cells(Rx + Rz + 2, 9 + Cz).Address())

    For Each iCells In tRange
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    Next iCells
    
    Pp.Range(Cells(Rx + 3, 10).Address(), Cells(Rx + Rz + 2, 9 + Cz).Address()).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
    Pp.Range(Cells(Rx + 3, 10).Address(), Cells(Rx + 3, 9 + Cz).Address()).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThick
        
        
    Pp.Columns("J:Z").HorizontalAlignment = xlCenter
    Pp.Columns("J:Z").AutoFit
    '--------------------------------------Changing Status in Main sheet
    ActiveWorkbook.Sheets("Main").Activate
    If Cz = 17 Then
        MsgBox "This is also a waveform"
        ActiveWorkbook.Sheets("Main").Activate
        Exit Sub
    ElseIf ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Wave Only" And Rx <> 0 Then
        ActiveWorkbook.Sheets("Main").Range(StatAdd).Value = "Stored"
    End If
    

End Sub
Public Sub Clears()
    Dim StatsAdd As String
    If Intersect(ActiveCell, Range("C20:C35")) Is Nothing Then
            MsgBox "Select the Meter Name to proceed. Try Again."
            Exit Sub
    Else
        If ActiveCell.Offset(, 2) = "" Then
            'MsgBox "Nothing to Clear"
            Exit Sub
        End If
    End If

 StatsAdd = ActiveCell.Offset(, 2).Address()
 For Each shts In ThisWorkbook.Worksheets 'Check if the worksheet is already existing
    If shts.Name = ActiveCell.Value Then
        Set Compa = ActiveWorkbook.Sheets(ActiveCell.Value)
        
        Compa.Activate
        Compa.UsedRange.ClearContents
        Compa.UsedRange.ClearFormats
        Cleardata = False
        Exit For
    End If
 Next shts
    
    
    ActiveWorkbook.Sheets("Main").Activate
    ActiveWorkbook.Sheets("Main").Range(StatsAdd).Value = ""
    ActiveWorkbook.Sheets("Main").Range(StatsAdd).Offset(, 1).Value = ""
    
End Sub
