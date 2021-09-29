Attribute VB_Name = "PrepAndCheck"
Sub PrepWorkbook()
    
    
    ActiveSheet.Name = "Cyber"
    Sheets.Add.Name = "GEARS"
    Sheets.Add.Name = "SDAP"
    Sheets.Add.Name = "DIAMOND"
    Sheets.Add.Name = "PML"
    Sheets.Add.Name = "ExactMatch"
    Sheets.Add.Name = "Dev"
    Sheets.Add.Name = "ServerSWUpdates"
    Sheets.Add.Name = "CompareNSLookupServers"
    Sheets.Add.Name = "CyberListfromCNSLookupS"
    Sheets.Add.Name = "Server"
    
   
    Worksheets("Cyber").Range("B:O").EntireColumn.Insert
    Worksheets("Cyber").Cells(1, 2) = "Label"
    Worksheets("Cyber").Cells(1, 3) = "SDAP EXACT MATCH"
    Worksheets("Cyber").Cells(1, 4) = "GEARS EXACT MATCH"
    Worksheets("Cyber").Cells(1, 5) = "Component Manual"
    Worksheets("Cyber").Cells(1, 6) = "Component from PML"
    Worksheets("Cyber").Cells(1, 7) = "Component from GEARS"
    Worksheets("Cyber").Cells(1, 8) = "LOGIC Match"
    Worksheets("Cyber").Cells(1, 9) = "Diamond lookup"
    Worksheets("Cyber").Cells(1, 10) = "Final Combined"
    Worksheets("Cyber").Cells(1, 11) = "Software Name"
    Worksheets("Cyber").Cells(1, 12) = "Software ID"
    Worksheets("Cyber").Cells(1, 13) = "Database"
    Worksheets("Cyber").Cells(1, 14) = "Server"
    Worksheets("Cyber").Cells(1, 15) = "NSLookup"
    
    Worksheets("GEARS").Range("A:H").EntireColumn.Insert
    Worksheets("GEARS").Cells(1, 1) = "ComponentLookup"
    Worksheets("GEARS").Cells(1, 2) = "ComponentLookup:ComponentID"
    Worksheets("GEARS").Cells(1, 3) = "ProductLookup"
    Worksheets("GEARS").Cells(1, 4) = "ProductLookup:EXTERNALID"
    Worksheets("GEARS").Cells(1, 5) = "Server"
    Worksheets("GEARS").Cells(1, 6) = "ServerExternalID"
    Worksheets("GEARS").Cells(1, 7) = "DatabaseInstance"
    Worksheets("GEARS").Cells(1, 8) = "NSLOOKUP"
    
    Worksheets("ServerSWUpdates").Range("A:K").EntireColumn.Insert
    Worksheets("ServerSWUpdates").Cells(1, 1) = "CyberServer"
    Worksheets("ServerSWUpdates").Cells(1, 2) = "Component"
    Worksheets("ServerSWUpdates").Cells(1, 3) = "Software"
    Worksheets("ServerSWUpdates").Cells(1, 4) = "ServerSoftware"
    Worksheets("ServerSWUpdates").Cells(1, 5) = "ServerID"
    Worksheets("ServerSWUpdates").Cells(1, 6) = "BigFix (ECMO)"
    Worksheets("ServerSWUpdates").Cells(1, 7) = "Yes"
    Worksheets("ServerSWUpdates").Cells(1, 8) = "SoftwareID"
    Worksheets("ServerSWUpdates").Cells(1, 9) = "Database"
    Worksheets("ServerSWUpdates").Cells(1, 10) = "GEARSEquivalent"
    Worksheets("ServerSWUpdates").Cells(1, 11) = "NSLookup"
    
    Worksheets("CompareNSLookupServers").Range("A:G").EntireColumn.Insert
    Worksheets("CompareNSLookupServers").Cells(1, 1) = "Component"
    Worksheets("CompareNSLookupServers").Cells(1, 2) = "NotFoundGEARS"
    Worksheets("CompareNSLookupServers").Cells(1, 3) = "FoundCyber"
    Worksheets("CompareNSLookupServers").Cells(1, 4) = "GEARS Servers"
    Worksheets("CompareNSLookupServers").Cells(1, 5) = "Cyber Servers"
    Worksheets("CompareNSLookupServers").Cells(1, 6) = "Lifecycle"
    Worksheets("CompareNSLookupServers").Cells(1, 7) = "Notes"
    
    
    Worksheets("CyberListfromCNSLookupS").Range("A:G").EntireColumn.Insert
    Worksheets("CyberListfromCNSLookupS").Cells(1, 1) = "Server"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 2) = "Component"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 3) = "Software"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 4) = "ServerSoftware"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 5) = "Server ID"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 6) = "BigFix (ECMO)"
    Worksheets("CyberListfromCNSLookupS").Cells(1, 7) = "Yes"

   
End Sub
Sub MatchOnSDAP()

Dim CyberRowCnt As Long
Dim SDAPRowCnt As Long
Dim CyberEnd As Long
Dim SDAPEnd As Long
Dim CyberComputerName As String
Dim SDAPComponentName As String
Dim SDAPServerName As String


Application.ScreenUpdating = False

SDAPServerNameCol = 4
SDAPCompNameCol = 6
CyberComputerNameCol = 1
CyberComponentNameFromSDAPCol = 3

CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
SDAPEnd = Worksheets("SDAP").Cells(Rows.Count, 1).End(xlUp).Row

Dim logline As String
Txt_Append "MatchOnSDAP - START"
For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 50 = 0 Then
        logline = "MatchOnSDAP| " & CyberRowCnt
        Txt_Append (logline)
    End If

    CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
    
    
    For SDAPRowCnt = 2 To SDAPEnd
    
     SDAPServerName = UCase(RTrim(LTrim(Worksheets("SDAP").Cells(SDAPRowCnt, SDAPServerNameCol).Value)))
      SDAPComponentName = UCase(RTrim(LTrim(Worksheets("SDAP").Cells(SDAPRowCnt, SDAPCompNameCol).Value)))

      If SDAPServerName = CyberComputerName Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromSDAPCol).Value = SDAPComponentName
     End If
    
   ' If SDAPRowCnt > 1654 Then
    '    x = 1
   ' End If
     
    Next SDAPRowCnt
    
   
  
Next CyberRowCnt
    Txt_Append "MatchOnSDAP - COMPLETE"

End Sub
Sub ExactCyberToGEARSServerMatch()
'Checks for exact server Match updates the GEARS COlumn

Dim ComponentNameGEARS As String
Dim CyberServerNameChk As String
Dim GEARSServerNameChk As String
Dim ServerMatchCnt As Integer
Dim CyberRowCnt As Integer
Dim GEARSRowCnt As Integer
Dim CyberEnd As Long
Dim GEARSEnd As Long

Application.ScreenUpdating = False

GEARSServerNameCol = 5
CyberComputerNameCol = 1
CyberComponentNameFromGEARSCol = 4

CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
GEARSEnd = Worksheets("GEARS").Cells(Rows.Count, 1).End(xlUp).Row

' Loop through Orginal ECMO/Cyber worksheet with added columns for Component identification

Dim logline As String
Txt_Append "ExactCyberToGEARSServerMatch - START"
For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 50 = 0 Then
        logline = "ExactCyberToGEARSServerMatch| " & CyberRowCnt
        Txt_Append (logline)
    End If
    
    CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
        

    
    For GEARSRowCnt = 2 To GEARSEnd
     
      GEARSServerName = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNameCol).Value)))

      If GEARSServerName = CyberComputerName Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromGEARSCol).Value = GEARSServerName
     End If
    Next GEARSRowCnt
  
Next CyberRowCnt

Txt_Append "ExactCyberToGEARSServerMatch - COMPLETE"

End Sub
Sub CheckManual()

Dim CyberRowCnt As Long
Dim ManualRowCnt As Long
Dim CyberEnd As Long
Dim ManualEnd As Long
Dim CyberComputerName As String
Dim ManualServerName As String
Dim ManualComponentName As String


Application.ScreenUpdating = False

ManualServerNameCol = 1
ManualCompNameCol = 2
CyberComputerNameCol = 1
CyberComponentNameFromManualCol = 5

CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
ManualEnd = Worksheets("Manual").Cells(Rows.Count, 1).End(xlUp).Row
Dim logline As String
Txt_Append "CheckManual - START"
For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 50 = 0 Then
        logline = "CheckManual| " & CyberRowCnt
        Txt_Append (logline)
    End If
    
    CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
    
    
    For ManualRowCnt = 2 To ManualEnd
    
     ManualServerName = UCase(RTrim(LTrim(Worksheets("Manual").Cells(ManualRowCnt, ManualServerNameCol).Value)))
      ManualComponentName = UCase(RTrim(LTrim(Worksheets("Manual").Cells(ManualRowCnt, ManualCompNameCol).Value)))

      If ManualServerName = CyberComputerName Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromManualCol).Value = ManualComponentName
     End If
    
         
    Next ManualRowCnt
    
   
  
Next CyberRowCnt
    Txt_Append "CheckManual - COMPLETE"

End Sub
Sub IdentifyType()

Dim i As Long
Dim CyberComputerNameCol As Integer
Dim CyberComputerName As String
Dim CyberRelayCol As Integer
Dim CyberRelay As String
Dim CyberReasonCol As Integer
Dim CyberOS As String
Dim CyberEMCol As Integer
Dim CyberExactMatch As String
Dim CyberEnd As Long
Dim CyberDateCol As Integer
Dim CyberDate As String
Dim TodayDate As String
Dim YesterdayDate As String
Dim DateFlag As Integer


Application.ScreenUpdating = False

DateFlag = 0
CyberComputerNameCol = 1
CyberReasonCol = 2
CyberEMCol = 4
CyberRelayCol = 26
CyberOSCol = 30
CyberDateCol = 32

CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row

Dim logline As String
Txt_Append "IdentifyType - START"

For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 100 = 0 Then
        logline = "IdentifyType| " & CyberRowCnt
        Txt_Append (logline)
    End If
    
    CyberComputerName = LCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
    CyberRelay = LCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberRelayCol).Value)))
    CyberOS = LCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberOSCol).Value)))
    CyberExactMatch = LCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberEMCol).Value)))
    CyberDate = (Worksheets("Cyber").Cells(CyberRowCnt, CyberDateCol).Value)

    
    ' TodayDate = Format(Date, "dd mmm yyyy")
    TodayDate = Format("9/27/2021", "dd mmm yyyy")
    

    'YesterdayDate = DateAdd("d", -1, Date)
    'YesterdayDate = Format(YesterdayDate, "dd mmm yyyy")
    
    Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = ""
    
    Select Case True
            
    Case InStr(CyberRelay, "dev-") > 0
        Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = "dev-dev"
   
    Case InStr(CyberOS, "win10") > 0
        Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = "dev-os"
        
    Case InStr(CyberDate, TodayDate) = 0
        Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = "dev-date"
        
    Case InStr(CyberExactMatch, "") > 0
        Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = "Exact Match"
        
    
    Case Else
         
    End Select
    
    
  
    
  '  If InStr(CyberDate, TodayDate) = 0 Then
   '    If InStr(CyberDate, YesterdayDate) > 0 Then
  '    Else
  '          Worksheets("Cyber").Cells(CyberRowCnt, CyberReasonCol).Value = "dev"
   '    End If
   ' End If
Next CyberRowCnt

Txt_Append "IdentifyType - COMPLETE"

End Sub
Sub MoveCyberNonProd()

    Dim cyberRange As Range
    Dim cyberRow As Long
    Dim devRow As Long
    Dim cyberRowCtr As Long
    Dim exactMatchRow As Long
    
    cyberRow = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
    devRow = Worksheets("Dev").Cells(Rows.Count, 1).End(xlUp).Row
    exactMatchRow = Worksheets("ExactMatch").Cells(Rows.Count, 1).End(xlUp).Row
    
    If devRow = 1 Then
       If Application.WorksheetFunction.CountA(Worksheets("Dev").UsedRange) = 0 Then devRow = 0
    End If
    
    If exactMatchRow = 1 Then
       If Application.WorksheetFunction.CountA(Worksheets("ExactMatch").UsedRange) = 0 Then exactMatchRow = 0
    End If
    
    Set cyberRange = Worksheets("Cyber").Range("B2:B" & cyberRow)
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    Dim logline As String
    Txt_Append "MoveCyberNonProd - START"
    For CyberRowCnt = 1 To cyberRange.Count
    
        If CyberRowCnt Mod 50 = 0 Then
            logline = "MoveCyberNonProd| " & CyberRowCnt
            Txt_Append (logline)
        End If
        
        If CStr(cyberRange(CyberRowCnt).Value) Like "dev-*" Then
            cyberRange(CyberRowCnt).EntireRow.Copy Destination:=Worksheets("Dev").Range("A" & devRow + 1)
            cyberRange(CyberRowCnt).EntireRow.Delete
            If CStr(cyberRange(CyberRowCnt).Value) Like "dev-*" Then
                CyberRowCnt = CyberRowCnt - 1
            End If
            devRow = devRow + 1
            If devRow Mod 50 = 0 Then
                logline = "MoveCyberNonProd|dev| " & devRow
                Txt_Append (logline)
            End If
        End If
        
        If CStr(cyberRange(CyberRowCnt).Value) = "Exact Match" Then
            cyberRange(CyberRowCnt).EntireRow.Copy Destination:=Worksheets("ExactMatch").Range("A" & exactMatchRow + 1)
            cyberRange(CyberRowCnt).EntireRow.Delete
            If CStr(cyberRange(CyberRowCnt).Value) = "Exact Match" Then
                CyberRowCnt = CyberRowCnt - 1
            End If
            exactMatchRow = exactMatchRow + 1
            If exactMatchRow Mod 50 = 0 Then
                logline = "MoveCyberNonProd|exactmatch| " & exactMatchRow
                Txt_Append (logline)
            End If
        End If
        
    Next
 Application.ScreenUpdating = True
    Txt_Append "MoveCyberNonProd - COMPLETE"
 End Sub

Sub DiamondCheck()

Dim CyberServerName As String
Dim DiamondServerName As String
Dim DiamondComponentName As String

Dim CyberRowCnt As Integer
Dim DiamondRowCnt As Integer
Dim DiamondMatchCol As Integer
Dim DiamondComponentCol As Integer

Dim CyberEnd As Long
Dim DiamondEnd As Long

Application.ScreenUpdating = False

CyberServerNameCol = 1
DiamondMatchCol = 9

DiamondServerNameCol = 1
DiamondComponentCol = 5

CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
DiamondEnd = Worksheets("Diamond").Cells(Rows.Count, 1).End(xlUp).Row
Dim logline As String
Txt_Append "DiamondCheck - START"

For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 50 = 0 Then
        logline = "DiamondCheck| " & CyberRowCnt
        Txt_Append (logline)
    End If

    CyberServerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberServerNameCol).Value)))

        For DiamondRowCnt = 2 To DiamondEnd
         
            DiamondServerName = UCase(RTrim(LTrim(Worksheets("Diamond").Cells(DiamondRowCnt, DiamondServerNameCol).Value)))
            DiamondServerName = LTrim(Replace(DiamondServerName, Chr(160), Chr(32)))
            DiamondComponentName = UCase(RTrim(LTrim(Worksheets("Diamond").Cells(DiamondRowCnt, DiamondComponentCol).Value)))

            If CyberServerName = DiamondServerName Then
            Worksheets("Cyber").Cells(CyberRowCnt, DiamondMatchCol).Value = DiamondComponentName
          
            End If
           
        Next DiamondRowCnt

   
Next CyberRowCnt
    Txt_Append "DiamondCheck - COMPLETE"

End Sub


Function Txt_Append(sText As String)
    On Error GoTo Err_Handler
    Dim iFileNumber           As Integer
 
    iFileNumber = FreeFile                   ' Get unused file number
    Dim sfFile As String
    sFile = "D:\source\ptoNodeProject\ptoNodeSP\data\test\output\vba.log"
  
    Open sFile For Append As #iFileNumber    ' Connect to the file
    Dim ts As String
    ts = Format(Now, "hhmmss") & Right(Format(Timer, "0.000"), 4)
    sText = ts & "|" & sText
    Print #iFileNumber, sText                ' Append our string
    Close #iFileNumber                       ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Txt_Append" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function
