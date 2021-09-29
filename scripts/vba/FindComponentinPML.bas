Attribute VB_Name = "FindComponentinPML"
Sub FindComponentinPML()
'fuzzy logic to find component to server association
'does delimited server name to component exact match
'does server pattern match using archserversoftware and archdatabase

Dim ComponentNameAry As Variant
Dim ComponentName As String
Dim i As Integer
Dim FoundComponentGreen As Integer
Dim PMLRowCnt As Long
Dim CyberRowCnt As Long
Dim ComponentNameCyber As String
Dim ComponentNamePML As String
Dim ComponentNamePMLAry As Variant
Dim iPML As Integer
Dim ConcatComponentNamePML As String
Dim HeaderRow As Integer
Dim GreenCnt As Long
Dim CommentValue As String
Dim CommentValueIn As String
Dim CyberServerNameChk As String
Dim GEARSServerNameChk As String
Dim ServerMatchCnt As Long
Dim SelectedGreenCnt  As Long
Dim ConcatenationFound As Boolean
Dim CyberServerAllAry() As Variant
Dim PMLComponentNameMatchAllAry(0 To 800) As String
Dim PMLComponentNameAllAry(0 To 800) As String
Dim PMLComponentNameAllCnt As Long
Dim CyberComputerNameLong As String
Dim CyberEnd As Long
Dim PMLEnd As Long
Dim GEARSSoftwareNameNoVersionCol As Integer
Dim CyberGEARSServerCol As Integer
Dim CyberGEARSServerNSLookupCol As Integer
Dim CyberComponentNameCombinedCol As Integer
Dim CyberSDAPComponentCol As Integer
Dim CyberDiamondComponentCol As Integer

ServerMatchCnt = 0
GreenCnt = 0
SelectedGreenCnt = 0
 

Application.ScreenUpdating = False

PMLComponentNameCol = 7
GEARSSoftwareNameCol = 3
GEARSSoftwareIDCol = 4
GEARSDatabaseCol = 7
GEARSSoftwareNameNoVersionCol = 14

CyberComputerNameCol = 1
CyberSDAPComponentCol = 3
CyberComponentNameManualCol = 5
CyberComponentNameFromPMLCol = 6
CyberComponentNameFromGEARSCol = 7
CyberComponentServerFuzzyLogicMatchCol = 8
CyberDiamondComponentCol = 9

CyberComponentNameCombinedCol = 10
CyberSoftwareNameCol = 11
CyberSoftwareIDCol = 12
CyberDatabaseCol = 13
CyberGEARSServerCol = 14

CyberGEARSServerNSLookupCol = 15

'Find Last Row in worksheets
CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
PMLEnd = Worksheets("PML").Cells(Rows.Count, 1).End(xlUp).Row

'CyberEnd = 100


HeaderRow = 1
Dim logline As String
' Delete old data
For CyberRowCnt = 2 To CyberEnd

    If CyberRowCnt Mod 50 = 0 Then
        logline = "FindCOmponentinPML|ClearColumns| " & CyberRowCnt
        Txt_Append (logline)
    End If
  Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromPMLCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromGEARSCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareNameCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareIDCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberDatabaseCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberGEARSServerCol).Value = ""
  Worksheets("Cyber").Cells(CyberRowCnt, CyberGEARSServerNSLookupCol).Value = ""
         
Next CyberRowCnt



' Creating a Array of PML



For PMLRowCnt = 2 To PMLEnd

    If PMLRowCnt Mod 50 = 0 Then
        logline = "PMLRowCnt|BuildArrays| " & PMLRowCnt
        Txt_Append (logline)
    End If
        ComponentNamePMLStart = UCase(RTrim(LTrim(Worksheets("PML").Cells(PMLRowCnt, PMLComponentNameCol).Value)))
        ComponentNamePML = WorksheetFunction.Substitute(ComponentNamePMLStart, " ", "-")
        ComponentNamePMLAry = Split(ComponentNamePML, "-")

        'Look for delimited Component Names Only - Reduces error by separating name parts that are significant and insignificant

        For iPML = 0 To UBound(ComponentNamePMLAry)

          ComponentNamePML = ComponentNamePMLAry(iPML)
          
          If Len(ComponentNamePML) > 1 And ComponentNamePML <> "WEB" And ComponentNameCyber <> "SERVICE" And ComponentNamePML <> "INFRA" And ComponentNamePML <> "DB" And ComponentNamePML <> "APP" And ComponentNamePML <> "USPTO" And ComponentNamePML <> "MGMT" Then
            PMLComponentNameMatchAllAry(PMLComponentNameAllCnt) = ComponentNamePML
            PMLComponentNameAllAry(PMLComponentNameAllCnt) = ComponentNamePMLStart
            
            If InStr(ComponentNamePMLStart, "-") > 0 Or InStr(ComponentNamePMLStart, " ") > 0 Then
              For i = 0 To UBound(PMLComponentNameMatchAllAry)
                 If ComponentNamePML = PMLComponentNameMatchAllAry(i) And ComponentNamePMLStart <> PMLComponentNameAllAry(i) Then
                    PMLComponentNameMatchAllAry(PMLComponentNameAllCnt) = ComponentNamePMLStart
                    PMLComponentNameMatchAllAry(PMLComponentNameAllCnt) = WorksheetFunction.Substitute(ComponentNamePMLStart, " ", "-")
                    
                 End If
              Next i
            End If
            PMLComponentNameAllCnt = PMLComponentNameAllCnt + 1
          End If
        Next iPML
Next PMLRowCnt

For iar = 0 To UBound(PMLComponentNameMatchAllAry)
    logline = "[" & PMLComponentNameMatchAllAry(iar) & "] & [" & PMLComponentNameAllAry(iar) & "]"
    PrepAndCheck.Txt_Append logline
Next iar

' Loop through Orginal ECMO/Cyber worksheet with added columns for Component identification

For CyberRowCnt = 2 To CyberEnd
    If CyberRowCnt Mod 50 = 0 Then
        logline = "FindCOmponentinPML|ComponentID| " & CyberRowCnt
        Txt_Append (logline)
    End If

   CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))

  If Left(CyberComputerName, 2) = "W-" Then
      CyberComputerName = Mid(CyberComputerName, 3, Len(CyberComputerName))
  End If

 

' Check Patterns of Servers

 
  CyberComputerNameHold = CyberComputerName

  ServerCheck CyberComputerNameHold, CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameFromGEARSCol, ServerMatchCnt, GEARSSoftwareNameCol, GEARSSoftwareIDCol, CyberSoftwareNameCol, CyberSoftwareIDCol, GEARSDatabaseCol, CyberDatabaseCol, GEARSSoftwareNameNoVersionCol, CyberGEARSServerCol, CyberGEARSServerNSLookupCol
' ServerCheck looks for exact match of the pattern and a string search of the pattern.  If exact match, it picks up the software otherwise it does not
' TODO: is the string search accurate enough to pickup software too
  
  
  If Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromGEARSCol).Value <> "" Then
  
  Else

   If InStr(CyberComputerName, "-") < 1 Then
    FoundComponentGreen = 99
   Else

   Worksheets("PML").Activate

   ComponentNameAry = Split(CyberComputerName, "-")
   CyberComputerNameLong = WorksheetFunction.Substitute(CyberComputerName, "-", "")
   CyberComputerNameLong = WorksheetFunction.Substitute(CyberComputerNameLong, ".", "")
   
   FoundComponentGreen = 99

   For i = 0 To UBound(ComponentNameAry)

     
        'Look for delimited Component Names Only - Reduces error by separating name parts that are significant and insignificant

        For iPML = 0 To UBound(PMLComponentNameMatchAllAry)
          ComponentNameOut = ""
          ComponentNamePML = PMLComponentNameMatchAllAry(iPML)
          
          ComponentNameCyber = ComponentNameAry(i)
          RemoveNumbers ComponentNameCyber
          RemoveNumbers CyberComputerNameLong
          
          
        If ComponentNamePML = CyberComputerNameLong Then
            FoundComponentGreen = i

            ComponentNameOut = CyberComputerNameLong
            iPML = UBound(PMLComponentNameMatchAllAry) + 1
            i = UBound(ComponentNameAry) + 1
        Else
        If Len(ComponentNameCyber) > 1 And ComponentNameCyber <> "WEB" And ComponentNameCyber <> "SERVICE" And ComponentNameCyber <> "INFRA" And ComponentNameCyber <> "DB" And ComponentNameCyber <> "APP" Then
              If ComponentNamePML = ComponentNameCyber Then
                 FoundComponentGreen = i

                 ComponentNameOut = PMLComponentNameAllAry(iPML)
                 
                 iPML = UBound(PMLComponentNameMatchAllAry) + 1
                 i = UBound(ComponentNameAry) + 1
                 
              End If
            
             
             End If
            
          End If
     
        Next iPML
   Next i
  End If
  
End If

Worksheets("Cyber").Activate

If FoundComponentGreen = 99 Then

Else

  Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromPMLCol).Value = ComponentNameOut


End If

  
Next CyberRowCnt

' Second Pass to find hyphenated

CorrectPMLwithDash CyberEnd, CyberComputerName, iPML, CyberComputerNameCol, ComponentNamePML, PMLComponentNameCol, CyberRowCnt, PMLRowCnt, ComponentNamePMLStart, PMLComponentNameAllCnt, PMLComponentNameMatchAllAry, CyberComponentNameFromGEARSCol, CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameFromPMLCol, CyberComponentNameManualCol
CommentValueIn = ""

CommentValue = ""

CommentValue = Str(GreenCnt) & "*" & "Green = Exact Match in '-' Seperated "
CommentValueIn = ""

CommentValue = ""

 

 'Add ServerMatchCnt Header

CommentValue = Str(ServerMatchCnt) & "* " & "Component that matches Servers in GEARS"

 

CommentValue = CommentValue & vbNewLine & SelectedGreenCnt & " " & "are Match exact"


For CyberRowCnt = 2 To CyberEnd
    FieldsEqualChk CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameCombinedCol, CyberComponentNameFromPMLCol, CyberComponentNameFromGEARSCol, CyberRowCnt, CyberComponentNameManualCol, SelectedGreenCnt
   
    Dim CyberComponentSDAP As String
    Dim CyberComponentDiamond As String
    Dim CyberComponentManual As String
    Dim CyberFuzzyLogic As String
    
    CyberComponentSDAP = Worksheets("Cyber").Cells(CyberRowCnt, CyberSDAPComponentCol).Value
    CyberComponentDiamond = Worksheets("Cyber").Cells(CyberRowCnt, CyberDiamondComponentCol).Value
    CyberComponentManual = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameManualCol).Value
    CyberFuzzyLogic = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol).Value
    
    If UCase(LTrim(RTrim(CyberComponentSDAP))) <> "" Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value = CyberComponentSDAP
    Else
    If UCase(LTrim(RTrim(CyberComponentDiamond))) <> "" Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value = CyberComponentDiamond
    Else
    If UCase(LTrim(RTrim(CyberComponentManual))) <> "" Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value = CyberComponentManual
    Else
    If UCase(LTrim(RTrim(CyberFuzzyLogic))) <> "" Then
        Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value = CyberFuzzyLogic
    End If
    End If
    End If
    End If
    
Next CyberRowCnt

End Sub

Sub CheckConcatenation(ComponentNamePMLAry, iPML, ComponentNameAry, i, ConcatenationFound, ConcatComponentNamePML)
  Dim ConcatComponentNameCyber As String
  
  If iPML = UBound(ComponentNamePMLAry) Or i = UBound(ComponentNameAry) Then
  Else
    ConcatComponentNamePML = ComponentNamePMLAry(iPML) & ComponentNamePMLAry(iPML + 1)
    ConcatComponentNameCyber = ComponentNameAry(i + 1)
    RemoveNumbers ConcatComponentNameCyber
    
    If ConcatComponentNamePML = ComponentNameAry(i) Or ConcatComponentNamePML = ComponentNameAry(i) & ConcatComponentNameCyber Then
       ConcatenationFound = True
       ConcatComponentNamePML = ComponentNamePMLAry(iPML) & "-" & ComponentNamePMLAry(iPML + 1)
       
    End If
  End If
End Sub

Sub RemoveNumbers(ComponentNameCyber)
x = 1
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "0", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "1", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "2", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "3", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "4", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "5", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "6", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "7", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "8", "")
    ComponentNameCyber = Application.WorksheetFunction.Substitute(ComponentNameCyber, "9", "")
    x = 2
    
End Sub

Sub ServerCheck(CyberComputerNameHold, CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameFromGEARSCol, ServerMatchCnt, GEARSSoftwareNameCol, GEARSSoftwareIDCol, CyberSoftwareNameCol, CyberSoftwareIDCol, GEARSDatabaseCol, CyberDatabaseCol, GEARSSoftwareNameNoVersionCol, CyberGEARSServerCol, CyberGEARSServerNSLookupCol)

    Dim NumCyberComputerNameHold As Boolean
    Dim LenComputerName As Integer
    Dim CharCnt As Long
    Dim CyberRowCntHold As Long
    Dim SoftwareNameGEARS As String
    Dim SoftwareIDGEARS As String
    Dim DatabaseGEARS As String
    Dim FoundSoftware As Boolean
    Dim SoftwareNameGEARSAry As Variant
    Dim iSoftwareNameAry As Integer
    Dim FoundDatabase As Boolean
    Dim DatabaseGEARSAry As Variant
    Dim iDatabaseAry As Integer
    Dim GEARSServerNameCol As Integer
    Dim GEARSServerNSLookupCol As Integer
    Dim GEARSComponentNameCol As Integer
    Dim ServerNSLookupGEARS As String
    Dim GEARSEnd As Long
    Dim CyberFuzzyLogic As String
    
    
  GEARSEnd = Worksheets("GEARS").Cells(Rows.Count, 1).End(xlUp).Row
   
    CharCnt = 1

    CyberRowCntHold = CyberRowCnt

    LenComputerName = Len(CyberComputerNameHold)
    
    RemoveNumbers CyberComputerNameHold


    LenComputerName = Len(CyberComputerNameHold)

    If Right(CyberComputerNameHold, 1) = "-" Then

      CyberComputerNameHold = Left(CyberComputerNameHold, LenComputerName - 1)

    End If

    GEARSServerNameCol = 5
    GEARServerIDCol = 6
    GEARSServerNSLookupCol = 8

    GEARSComponentNameCol = 1

    Worksheets("GEARS").Activate
ServerNSLookupGEARS = ""
 For GEARSRowCnt = 2 To GEARSEnd

      ServerNameGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNameCol).Value)))
      'ServerIDGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerIDCol).Value)))
      ServerNSLookupGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNSLookupCol).Value)))
      SoftwareNameGEARS = RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSSoftwareNameCol).Value))
      SoftwareIDGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSSoftwareIDCol).Value)))
      DatabaseGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSDatabaseCol).Value)))

      ComponentNameGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value)))
      CyberFuzzyLogic = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(GEARSRowCnt, CyberComponentServerFuzzyLogicMatchCol).Value)))
      
      ServerNameGEARSHold = ServerNameGEARS
      RemoveNumbers ServerNameGEARS

'debug
If CyberRowCntHold > 200 Then
x = 1

End If
If GEARSRowCnt > 7065 Then
x = 1


End If

    LenComputerName = Len(ServerNameGEARS)

    If Right(ServerNameGEARS, 1) = "-" Then

      ServerNameGEARS = Left(ServerNameGEARS, LenComputerName - 1)

    End If
  
    If ServerNameGEARS = CyberComputerNameHold Then
     
         If Trim(Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareNameCol).Value) = "" Then
            Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareNameCol).Value = SoftwareNameGEARS
            Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareIDCol).Value = SoftwareIDGEARS
            Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value = DatabaseGEARS
            Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerCol).Value = ServerNameGEARSHold
           
            Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerNSLookupCol).Value = ServerNSLookupGEARS
         Else
            FoundSoftware = False
            SoftwareNameGEARSAry = Split(Trim(Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareNameCol).Value), ";")
            For iSoftwareNameAry = 0 To UBound(SoftwareNameGEARSAry)
              If SoftwareNameGEARS = SoftwareNameGEARSAry(iSoftwareNameAry) Then
                FoundSoftware = True
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerCol).Value & "," & ServerNameGEARSHold
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerNSLookupCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerNSLookupCol).Value & "," & ServerNSLookupGEARS
              End If
              
            Next iSoftwareNameAry
            
            If FoundSoftware = False Then
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareNameCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareNameCol).Value & ";" & SoftwareNameGEARS
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareIDCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberSoftwareIDCol).Value & ";" & SoftwareIDGEARS
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value & ";" & DatabaseGEARS
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerCol).Value & ";" & ServerNameGEARSHold
                Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerNSLookupCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberGEARSServerNSLookupCol).Value & ";" & ServerNSLookupGEARS
                
            End If
            
            'FoundDatabase = False
            'DatabaseGEARSAry = Split(Trim(Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value), ";")
            'For iDatabaseAry = 0 To UBound(DatabaseGEARSAry)
            '  If DatabaseGEARS = DatabaseGEARSAry(iDatabaseAry) Then
            '    FoundDatabase = True
                 
             ' End If
              
            'Next iDatabaseAry
            
            'If FoundDatabase = False And DatabaseGEARS <> "" Then
              'If Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value = "" Then
              '  Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value = DatabaseGEARS
              'Else
              '  Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value = Worksheets("Cyber").Cells(CyberRowCntHold, CyberDatabaseCol).Value & ";" & DatabaseGEARS
              'End If
            'End If
          End If
    Else
     
         ServerMatchCnt = ServerMatchCnt + 1

          'GEARSRowCnt = 7069

    End If
        
        ServerNameGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNameCol).Value)))
        ServerNSLookupGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNSLookupCol).Value)))

        ComponentNameGEARS = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value)))
    
        If InStr(ServerNameGEARS, CyberComputerNameHold) > 0 Then

          Worksheets("Cyber").Cells(CyberRowCntHold, CyberComponentNameFromGEARSCol).Value = ComponentNameGEARS
              

         ServerMatchCnt = ServerMatchCnt + 1

        'GEARSRowCnt = 7069

        End If


    Next GEARSRowCnt
   

End Sub

Sub FieldsEqualChk(CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameCombinedCol, CyberComponentNameFromPMLCol, CyberComponentNameFromGEARSCol, CyberRowCnt, CyberComponentNameManualCol, SelectedGreenCnt)

   Dim CyberComponentNameFromGEARS As String
   Dim CyberComponentNameFromPML As String
   Dim CyberComponentNameManual As String
   Dim CyberLogic As String


   CyberComponentNameFromGEARS = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromGEARSCol).Value
   CyberComponentNameFromPML = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromPMLCol).Value
   CyberComponentNameManual = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameManualCol).Value
   CyberComponentCombined = Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value

   If CyberRowCnt > 220 Then
   x = 1
   End If
   
   
   If UCase(LTrim(RTrim(CyberComponentNameFromGEARS))) <> "" Then
      Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol).Value = CyberComponentNameFromGEARS
      SelectedGreenCnt = SelectedGreenCnt + 1
   Else
   
   If UCase(LTrim(RTrim(CyberComponentNameFromPML))) <> "" Then
      Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentServerFuzzyLogicMatchCol).Value = CyberComponentNameFromPML
      SelectedGreenCnt = SelectedGreenCnt + 1
   End If

   End If
   

End Sub


Sub CorrectPMLwithDash(CyberEnd, CyberComputerName, iPML, CyberComputerNameCol, ComponentNamePML, PMLComponentNameCol, CyberRowCnt, PMLRowCnt, ComponentNamePMLStart, PMLComponentNameAllCnt, PMLComponentNameMatchAllAry, CyberComponentNameFromGEARSCol, CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameFromPMLCol, CyberComponentNameManualCol)

Dim PMLEnd As Long

PMLEnd = Worksheets("PML").Cells(Rows.Count, 1).End(xlUp).Row

PMLComponentNameAllCnt = 0

For PMLRowCnt = 2 To PMLEnd

        ComponentNamePMLStart = UCase(RTrim(LTrim(Worksheets("PML").Cells(PMLRowCnt, PMLComponentNameCol).Value)))
        
        PMLComponentNameMatchAllAry(PMLComponentNameAllCnt) = ComponentNamePMLStart
            
        PMLComponentNameAllCnt = PMLComponentNameAllCnt + 1
Next PMLRowCnt

For CyberRowCnt = 2 To 100


    ComponentNamePML = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromPMLCol).Value)))
    
    If InStr(ComponentNamePML, "-") > 0 Then
      CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
        
      If InStr(CyberComputerName, ComponentNamePML) = 0 Then
        For iPML = 0 To UBound(PMLComponentNameMatchAllAry)
          ComponentNamePML = PMLComponentNameMatchAllAry(iPML)
          If InStr(CyberComputerName, ComponentNamePML) > 0 And ComponentNamePML <> "" Then
             Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameFromPMLCol).Value = ComponentNamePML
             iPML = UBound(PMLComponentNameMatchAllAry)
          End If
        Next iPML
      End If
    End If
    
'    FieldsEqualChk CyberComponentServerFuzzyLogicMatchCol, CyberComponentNameFromPMLCol, CyberComponentNameFromGEARSCol, CyberRowCnt, CyberComponentNameManualCol, SelectedGreenCnt

Next CyberRowCnt

End Sub
