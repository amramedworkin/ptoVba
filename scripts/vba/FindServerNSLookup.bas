Attribute VB_Name = "FindServerNSLookup"
Sub FindServersNSLookup()


Dim ComponentNameAry As Variant
Dim CyberServerListAry As Variant
Dim ComponentName As String
Dim i As Integer
Dim FoundServer As Boolean
Dim iCyberServerList As Integer
Dim FoundComponentGreen As Integer
Dim GEARSRowCnt As Long
Dim CyberRowCnt As Long
Dim ComponentNameCyber As String
Dim ComponentNamePML As String
Dim ComponentNamePMLAry As Variant
Dim iPML As Integer
Dim ConcatComponentNamePML As String
Dim HeaderRow As Integer
Dim GreenCnt As Integer
Dim CommentValue As String
Dim CommentValueIn As String
Dim CyberServerNameChk As String
Dim GEARSServerNameChk As String
Dim ServerMatchCnt As Long
Dim SelectedGreenCnt  As Long
Dim ConcatenationFound As Boolean
Dim CyberServerAllAry() As Variant
Dim CyberComputerNameLong As String
Dim GEARSEnd As Long
Dim CyberEnd As Long
Dim GEARSSoftwareNameNoVersionCol As Integer
Dim CyberGEARSServerCol As Integer
Dim CyberGEARSServerNSLookupCol As Integer
Dim CyberComponentNameCombinedCol As Integer
Dim CompareNSLookupServersComponentNameCol As Integer
Dim CompareNSLookupServersNotFoundCntCol As Integer
Dim CompareNSLookupServersCyberCntCol As Integer
Dim CompareNSLookupServersGEARSServersCol As Integer
Dim CompareNSLookupServersCyberServersCol As Integer
Dim CyberComponentServerCnt As Long
Dim CyberComputerNameCol As Long
Dim CyberComputerName As String
Dim CyberListfromCNSLookupSRowCnt As Integer
Dim CyberListfromCNSLookupSServersCyberServersCol As Integer
Dim FoundComponent As Boolean

Dim CyberListfromCNSLookupSServersCyberSoftwareCol As Integer
Dim CyberListfromCNSLookupSServersCyberServerSoftwareCol As Integer
Dim CyberListfromCNSLookupSServersCyberServerIDCol As Integer
Dim CyberListfromCNSLookupSServersCyberBigFixCol As Integer
Dim CyberListfromCNSLookupSServersCyberYesCol As Integer

Dim ServerName As String
Dim ServerID As String
Dim ServerRowCnt As Long
Dim ServerEnd As Integer

ServerMatchCnt = 0
GreenCnt = 0
SelectedGreenCnt = 0
 

Application.ScreenUpdating = False


CompareNSLookupServersComponentNameCol = 1
CompareNSLookupServersGEARSNotFoundCntCol = 2
CompareNSLookupServersCyberCntCol = 3
CompareNSLookupServersGEARSServersCol = 4
CompareNSLookupServersCyberServersCol = 5
CompareNSLookupServersReplaceEqualCol = 6
'CompareNSLookupServersLifecycleCol = 6

GEARSComponentNameCol = 1
GEARSSoftwareNameCol = 3
GEARSSoftwareIDCol = 4
GEARSServerNameCol = 5
GEARSDatabaseCol = 7
GEARSNSLookupCol = 8
'GEARSLifecycleCol = 9

CyberComputerNameCol = 1
CyberComponentNameManualCol = 5
CyberComponentNameFromPMLCol = 6
'CyberComponentServerNameMatchCol = 7
'CyberComponentServerNameMatchExactCol = 8
CyberComponentNameCombinedCol = 10
CyberSoftwareNameCol = 11
CyberSoftwareIDCol = 12
CyberDatabaseCol = 13
CyberGEARSServerCol = 14
CyberGEARSServerNSLookupCol = 15

CyberListfromCNSLookupSServersCyberServersCol = 1
CyberListfromCNSLookupSServersCyberComponentNameCol = 2
CyberListfromCNSLookupSServersCyberSoftwareCol = 3
CyberListfromCNSLookupSServersCyberServerSoftwareCol = 4
CyberListfromCNSLookupSServersCyberServerIDCol = 5
CyberListfromCNSLookupSServersCyberBigFixCol = 6
CyberListfromCNSLookupSServersCyberYesCol = 7

ServerNameCol = 1
ServerIDCol = 2



'Find Last Row in worksheets
CompareNSLookupServersEnd = Worksheets("CompareNSLookupServers").Cells(Rows.Count, 1).End(xlUp).Row
CyberListfromCNSLookupSEnd = Worksheets("CyberListfromCNSLookupS").Cells(Rows.Count, 1).End(xlUp).Row
CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row
GEARSEnd = Worksheets("GEARS").Cells(Rows.Count, 6).End(xlUp).Row
ServerEnd = Worksheets("Server").Cells(Rows.Count, 1).End(xlUp).Row


HeaderRow = 1
' Delete old data
For CompareNSLookupServersRowCnt = 2 To CompareNSLookupServersEnd
  Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersComponentNameCol).Value = ""
  Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSNotFoundCntCol).Value = ""
  Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberCntCol).Value = ""
  Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSServersCol).Value = ""
  Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberServersCol).Value = ""
  'Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersLifecycleCol).Value = ""
Next CompareNSLookupServersRowCnt



For CyberListfromCNSLookupSRowCnt = 2 To CyberListfromCNSLookupSEnd
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberServersCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberComponentNameCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberSoftwareCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberServerSoftwareCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberServerIDCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberBigFixCol).Value = ""
  Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberYesCol).Value = ""
Next CyberListfromCNSLookupSRowCnt

'sort on component
Set GEARSSheet = ActiveWorkbook.Worksheets("GEARS")
Worksheets("GEARS").Select
Columns("A:L").Sort key1:=Range("A2"), order1:=xlAscending, Header:=xlYes

'sort on combined column
Set CyberSheet = ActiveWorkbook.Worksheets("Cyber")
Worksheets("Cyber").Select
Columns("A:AE").Sort key1:=Range("J2"), order1:=xlAscending, Header:=xlYes
   
'---------GEARS

CompareNSLookupServersRowCnt = 2

For GEARSRowCnt = 2 To GEARSEnd

    
   GEARSComponentName = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value)))
   GEARSComponentNameHold = GEARSComponentName
   GEARSServerName = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNameCol).Value)))
   GEARSNSLookup = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSNSLookupCol).Value)))
   'GEARSLifecycle = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSLifecycleCol).Value)))
     
   GEARSComponentServerCnt = 0
   GEARSServerNameHold = ""
   
   Do While GEARSComponentName = GEARSComponentNameHold
    
    
    If GEARSNSLookup = "NOTFOUND" Then
        GEARSServerNameHold = GEARSServerNameHold & ";" & GEARSServerName
        GEARSComponentServerCnt = GEARSComponentServerCnt + 1
    End If
    GEARSRowCnt = GEARSRowCnt + 1
    GEARSServerName = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerNameCol).Value)))
    GEARSComponentName = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value)))
    GEARSNSLookup = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSNSLookupCol).Value)))
    'GEARSLifecycle = UCase(RTrim(LTrim(Worksheets("GEARS").Cells(GEARSRowCnt, GEARSLifecycleCol).Value)))
    
    If GEARSRowCnt > GEARSEnd Then
      GEARSComponentName = ""
    End If
    
        
   Loop
   
   'Trim Leading";"
    
    Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersComponentNameCol).Value = GEARSComponentNameHold
    
    Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSServersCol).Value = Mid(GEARSServerNameHold, 2, Len(GEARSServerNameHold))
     
    Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSNotFoundCntCol).Value = GEARSComponentServerCnt - 1
    'Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersLifecycleCol).Value = GEARSLifecycle
        
    CompareNSLookupServersRowCnt = CompareNSLookupServersRowCnt + 1
Next GEARSRowCnt
     
'-------Cyber

CompareNSLookupServersRowCnt = 2
CyberListfromCNSLookupSRowCnt = 2
   
For CyberRowCnt = 2 To CyberEnd

    
   CyberComponentName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value)))
   CyberComponentNameHold = CyberComponentName
   CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
   CyberSoftwareName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareNameCol).Value)))
   CyberComponentServerCnt = 1
   CyberServerList = ""
     
   Do While CyberComponentName = CyberComponentNameHold And CyberComponentName <> ""
         ServerID = ""
        For ServerRowCnt = 2 To ServerEnd
            ServerName = UCase(RTrim(LTrim(Worksheets("Server").Cells(ServerRowCnt, ServerNameCol).Value)))
            If CyberComputerName = ServerName Then
                ServerID = UCase(RTrim(LTrim(Worksheets("Server").Cells(ServerRowCnt, ServerIDCol).Value)))
            End If
        Next ServerRowCnt
        
    If CyberSoftwareName = "" Then
        CyberComponentServerCnt = CyberComponentServerCnt + 1
        CyberServerListAry = Split(CyberServerList, ";")

        'Look for delimited Component Names Only - Reduces error by separating name parts that are significant and insignificant
        FoundServer = False
        
        For iCyberServerList = 0 To UBound(CyberServerListAry)
            If CyberServerListAry(iCyberServerList) = CyberComputerName Then
               FoundServer = True
            End If
        Next iCyberServerList
        
        If FoundServer = False Then
            CyberServerList = CyberServerList & ";" & CyberComputerName
            Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberServersCol).Value = LCase(CyberComputerName)
            Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberComponentNameCol).Value = CyberComponentName
            Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberServerIDCol).Value = ServerID
    Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberBigFixCol).Value = "BigFix (ECMO)"
    Worksheets("CyberListfromCNSLookupS").Cells(CyberListfromCNSLookupSRowCnt, CyberListfromCNSLookupSServersCyberYesCol).Value = "Yes"
            CyberListfromCNSLookupSRowCnt = CyberListfromCNSLookupSRowCnt + 1
        End If
    End If
    
    
   CyberComputerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComputerNameCol).Value)))
   CyberComponentName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCombinedCol).Value)))
   CyberSoftwareName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareNameCol).Value)))
   CyberRowCnt = CyberRowCnt + 1
    
    
    If CyberRowCnt > CyberEnd Then
      CyberComponentName = ""
    End If
   Loop
   
   'Trim Leading";"
    
    CyberServerList = Mid(CyberServerList, 2, Len(CyberServerList))
    FoundComponent = False
    
    For CompareNSLookupServersRowCnt = 2 To CompareNSLookupServersEnd
      If CyberComponentNameHold = UCase(RTrim(LTrim(Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersComponentNameCol).Value))) Then
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersComponentNameCol).Value = CyberComponentNameHold
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberCntCol).Value = CyberComponentServerCnt - 1
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberServersCol).Value = CyberServerList
        
        If CyberComponentServerCnt = Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSNotFoundCntCol).Value Then
          Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersReplaceEqualCol).Value = "All"
        End If
        FoundComponent = True
      End If
    Next CompareNSLookupServersRowCnt
    'CompareNSLookupServersRowCnt = CompareNSLookupServersRowCnt + 1
   If FoundComponent = False Then
        CompareNSLookupServersEnd = CompareNSLookupServersEnd + 1
        CompareNSLookupServersRowCnt = CompareNSLookupServersEnd
        
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersComponentNameCol).Value = CyberComponentNameHold
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberCntCol).Value = CyberComponentServerCnt - 1
        Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersCyberServersCol).Value = CyberServerList
        
        If CyberComponentServerCnt = Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersGEARSNotFoundCntCol).Value Then
          Worksheets("CompareNSLookupServers").Cells(CompareNSLookupServersRowCnt, CompareNSLookupServersReplaceEqualCol).Value = "All"
        End If
        
   End If
  
Next CyberRowCnt
  

'Worksheets("GEARS").Activate


End Sub






