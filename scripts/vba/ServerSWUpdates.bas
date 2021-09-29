Attribute VB_Name = "ServerSWUpdates"
Sub ServerSWUpdates()

Dim myrange As Range


Dim CyberServerName As String
Dim CyberComponentName As String
Dim CyberSoftwareName As String
Dim CyberSoftwareID As String
Dim ServerSWUpdatesName As String
Dim ServerSWUpdatesComponentName As String
Dim ServerSWUpdatesSoftwareName As String
Dim ServerSWUpdatesSoftwareID As String
Dim CyberRowCnt As Integer
Dim ServerSWUpdatesRowCnt As Integer

Dim CyberServerNameCol As Integer
Dim CyberComponentNameCol As Integer
Dim CyberSoftwareNameCol As Integer
Dim CyberSoftwareIDCol As Integer

Dim ServerSWUpdatesNameCol As Integer
Dim ServerSWUpdatesComponentCol As Integer
Dim ServerSWUpdatesSoftwareNameCol As Integer
Dim ServerSWUpdatesSoftwareIDCol As Integer
Dim SoftwareNameAry As Variant
Dim SoftwareIDAry As Variant
Dim DBAry As Variant
Dim CyberServerGEARSAry As Variant
Dim CyberServerGEARSNSLookupAry As Variant
Dim iSoftwareName As Integer
Dim GEARSComponentNameCol As Integer
Dim GEARSSoftwareIDCol As Integer
Dim GEARSServerCol As Integer
Dim GEARSNSLookupCol As Integer
Dim GEARSDatabaseCol As Integer
Dim GEARSServerSWUpdatesMarkedCol As Integer
Dim ServerSWUpdatesServerSoftwareCol As Integer
Dim ServerSWUpdatesBigFixCol As Integer
Dim ServerSWUpdatesYesCol As Integer



Dim CyberEnd As Long
Dim ServerSWUpdates As Long
Dim CyberServerGEARS As String
Dim CyberServerGEARSCol As Integer
Dim CyberServerGEARSNSLookup As String
Dim CyberServerGEARSNSLookupCol As Integer
Dim ServerSWUpdatesServerIDCol As Integer
Dim ServerRowCnt As Long
Dim ServerEnd As Long
Dim ServerNameCol As Long
Dim ServerIDCol As Long
Dim ServerName As String
Dim ServerID As String



Application.ScreenUpdating = False

GEARSComponentNameCol = 1
GEARSSoftwareIDCol = 4
GEARSServerCol = 5
GEARSDatabaseCol = 7
GEARSNSLookupCol = 8
GEARSServerSWUpdatesMarkedCol = 13

CyberServerNameCol = 1
CyberComponentNameCol = 10
CyberSoftwareNameCol = 11
CyberSoftwareIDCol = 12
CyberDBCol = 13
CyberServerGEARSCol = 14
CyberServerGEARSNSLookupCol = 15

ServerSWUpdatesNameCol = 1
ServerSWUpdatesComponentNameCol = 2
ServerSWUpdatesSoftwareNameCol = 3
ServerSWUpdatesServerSoftwareCol = 4
ServerSWUpdatesServerIDCol = 5
ServerSWUpdatesBigFixCol = 6
ServerSWUpdatesYesCol = 7
ServerSWUpdatesSoftwareIDCol = 8
ServerSWUpdatesDBCol = 9
ServerSWUpdatesServerGEARSCol = 10
ServerSWUpdatesServerGEARSNSLookupCol = 11

ServerNameCol = 1
ServerIDCol = 2


CyberEnd = Worksheets("Cyber").Cells(Rows.Count, 1).End(xlUp).Row + 1
ServerSWUpdatesEnd = Worksheets("ServerSWUpdates").Cells(Rows.Count, 1).End(xlUp).Row + 1
ServerEnd = Worksheets("Server").Cells(Rows.Count, 1).End(xlUp).Row + 1


    'If ServerSWUpdatesEnd = 1 Then
    '   If Application.WorksheetFunction.CountA(Worksheets("ServerSWUpdates").UsedRange) = 0 Then ServerSWUpdates = 0
    'End If
    
    'Set myrange = Worksheets("Cyber").Range("A2:A" & CyberEnd)
    
For ServerSWUpdatesRowCnt = 2 To ServerSWUpdatesEnd
  Worksheets("ServerSWUpdates").Range("A2").EntireRow.Delete
Next ServerSWUpdatesRowCnt
ServerSWUpdatesEnd = Worksheets("ServerSWUpdates").Cells(Rows.Count, 1).End(xlUp).Row + 1

For CyberRowCnt = 2 To CyberEnd
    CyberSoftwareName = RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareNameCol).Value))
    SoftwareNameAry = Split(CyberSoftwareName, ";")
     
     If CyberSoftwareName <> "" Then
    
      CyberServerName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberServerNameCol).Value)))
      CyberComponentName = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberComponentNameCol).Value)))
      CyberSoftwareID = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberSoftwareIDCol).Value)))
      CyberDB = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberDBCol).Value)))
      CyberServerGEARS = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberServerGEARSCol).Value)))
      CyberServerGEARSNSLookup = UCase(RTrim(LTrim(Worksheets("Cyber").Cells(CyberRowCnt, CyberServerGEARSNSLookupCol).Value)))

    ServerID = ""
      For ServerRowCnt = 2 To ServerEnd
       ServerName = UCase(RTrim(LTrim(Worksheets("Server").Cells(ServerRowCnt, ServerNameCol).Value)))
      
      
       If CyberServerName = ServerName Then
         ServerID = UCase(RTrim(LTrim(Worksheets("Server").Cells(ServerRowCnt, ServerIDCol).Value)))
       End If
            
       
      
      Next ServerRowCnt


      SoftwareIDAry = Split(CyberSoftwareID, ";")
      
      If CyberDB = "" Then
         CyberDB = ";"
      End If
      
   
      
         
      
      DBAry = Split(CyberDB, ";")
      CyberServerGEARSAry = Split(CyberServerGEARS, ";")
      CyberServerGEARSNSLookupAry = Split(CyberServerGEARSNSLookup, ";")
      
      For iSoftwareName = 0 To UBound(SoftwareNameAry)
      'If CyberRowCnt > 816 Then
      'x = 1
      'End If
        
        CyberServerName = LCase(CyberServerName)
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesNameCol).Value = CyberServerName
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesComponentNameCol).Value = CyberComponentName
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesSoftwareNameCol).Value = SoftwareNameAry(iSoftwareName)
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesServerSoftwareCol).Value = SoftwareNameAry(iSoftwareName) + " on " + CyberServerName
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesServerIDCol).Value = ServerID
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesBigFixCol).Value = "BigFix (ECMO)"
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesYesCol).Value = "Yes"
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesSoftwareIDCol).Value = SoftwareIDAry(iSoftwareName)
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesDBCol).Value = DBAry(iSoftwareName)
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesServerGEARSCol).Value = CyberServerGEARSAry(iSoftwareName)
        Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesEnd, ServerSWUpdatesServerGEARSNSLookupCol).Value = CyberServerGEARSNSLookupAry(iSoftwareName)
        

        ServerSWUpdatesEnd = ServerSWUpdatesEnd + 1
      Next iSoftwareName
      
      
    End If
    
Next CyberRowCnt

ResolveGEARSServersOneToOne ServerSWUpdatesNameCol, ServerSWUpdatesComponentNameCol, ServerSWUpdatesSoftwareIDCol, ServerSWUpdatesEnd, GEARSComponentNameCol, GEARSSoftwareIDCol, GEARSServerCol, GEARSNSLookupCol, GEARSServerSWUpdatesMarkedCol, ServerSWUpdatesServerGEARSCol

'FindDatabaseServerChanges ServerSWUpdatesComponentNameCol, ServerSWUpdatesSoftwareIDCol, ServerSWUpdatesEnd, GEARSComponentNameCol, GEARSDatabaseCol, GEARSServerCol, GEARSNSLookupCol, GEARSServerSWUpdatesMarkedCol

End Sub

Sub ResolveGEARSServersOneToOne(ServerSWUpdatesNameCol, ServerSWUpdatesComponentNameCol, ServerSWUpdatesSoftwareIDCol, ServerSWUpdatesEnd, GEARSComponentNameCol, GEARSSoftwareIDCol, GEARSServerCol, GEARSNSLookupCol, GEARSServerSWUpdatesMarkedCol, ServerSWUpdatesServerGEARSCol)
 'Find GEARS Components with equal servers to be replaced
 
   Dim ServerSWUpdatesComponentName As String
   Dim ServerSWUpdatesSoftwareID As String
   Dim ServerSWUpdatesServerName As String
   Dim GEARSComponentName As String
   Dim GEARSSoftwareID As String
   Dim ServerSWUpdatesRowCnt As Integer
   Dim GEARSRowCnt As Integer
   Dim GEARSLastRow As Integer
   Dim FoundServerSWUpdatesServerName As Boolean
   Dim iGEARSServerSWUpdatesMarked As Integer
   Dim GEARSServerSWUpdatesMarkedAry As Variant
   Dim GEARSServerSWUpdatesMarked As String
   Dim FoundServerSWUpdatesServerGEARS As Boolean
   Dim ServerSWUpdatesServerGEARS As String
   Dim iServerSWUpdatesServerGEARS As Integer
   Dim ServerSWUpdatesServerGEARSAry As Variant
   
   
   'GEARSLastRow = Worksheets("GEARS").Cells(Rows.Count, 1).End(xlUp).Row
   
   'For ServerSWUpdatesRowCnt = 1 To ServerSWUpdatesEnd
   '  ServerSWUpdatesComponentName = Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesComponentNameCol).Value
   '  ServerSWUpdatesSoftwareID = Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesSoftwareIDCol).Value
   '  ServerSWUpdatesServerName = Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesNameCol).Value
     
   '  For GEARSRowCnt = 2 To ServerSWUpdatesEnd
   '    GEASRComponentName = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value
   '    GEARSSoftwareID = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSSoftwareIDCol).Value
   '    GEARSServer = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerCol).Value
   '    GEARSNSLookup = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSNSLookupCol).Value
       
   '    If GEASRComponentName = ServerSWUpdatesComponentName And GEARSSoftwareID = GEARSSoftwareID Then
   '      If GEARSNSLookup = "NotFound" Then
   '        GEARSServerSWUpdatesMarked = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerSWUpdatesMarkedCol).Value
   '        If GEARSServerSWUpdatesMarked = "" Then
   '          Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerSWUpdatesMarkedCol).Value = ServerSWUpdatesServerName
   '        Else
   '          GEARSServerSWUpdatesMarkedAry = Split(GEARSServerSWUpdatesMarked, ";")
   '          FoundServerSWUpdatesServerName = False
   '          For iGEARSServerSWUpdatesMarked = 0 To UBound(GEARSServerSWUpdatesMarkedAry)
   '             If GEARSServerSWUpdatesMarkedAry(iGEARSServerSWUpdatesMarked) = ServerSWUpdatesServerName Then
   '               FoundServerSWUpdatesServerName = True
   '            End If
   '             Next iGEARSServerSWUpdatesMarked
   '             If FoundServerSWUpdatesServerName = False Then
   '               Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerSWUpdatesMarkedCol).Value = GEARSServerSWUpdatesMarked & ";" & ServerSWUpdatesServerName
   '             End If
   '        End If
           
           
           'ServerSWUpdatesServerGEARS = Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesServerGEARSCol).Value
           'ServerSWUpdatesServerGEARSAry = Split(ServerSWUpdatesServerGEARS, ";")
             
           'If ServerSWUpdatesServerGEARS = "" Then
           '  Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesServerGEARSCol).Value = GEARSServer
           'Else
            ' FoundServerSWUpdatesServerGEARS = False
             'For iServerSWUpdatesServerGEARS = 0 To ServerSWUpdatesServerGEARSAry
              ' If ServerSWUpdatesServerGEARSAry(iServerSWUpdatesServerGEARS) = ServerSWUpdatesServerGEARS Then
               '  FoundServerSWUpdatesServerGEARS = True
                'End If
             'Next iServerSWUpdatesServerGEARS
             'If FoundServerSWUpdatesServerGEARS = False Then
              '  Worksheets("ServerSWUpdates").Cells(ServerSWUpdatesRowCnt, ServerSWUpdatesServerGEARSCol).Value = ServerSWUpdatesServerGEARS & ";" & GEARSServer
             'End If
             
   '        End If
   '      End If
   '    End If
   '  Next GEARSRowCnt
     
   'Next ServerSWUpdatesRowCnt
   
End Sub


Sub FindDatabaseServerChanges(ServerSWUpdatesComponentNameCol, ServerSWUpdatesSoftwareIDCol, ServerSWUpdatesEnd, GEARSComponentNameCol, GEARSDatabaseCol, GEARSServerCol, GEARSNSLookupCol, GEARSServerSWUpdatesMarkedCol)
' Find Database servers not on the network and find Cyber equivelant to replace them

   Dim GEARSComponentName As String
   Dim GEARSDatabase As String
   Dim ServerSWUpdatesRowCnt As Integer
   Dim GEARSRowCnt As Integer
   Dim GEARSLastRow As Integer
   
   For GEARSRowCnt = 2 To ServerSWUpdates
       GEASRComponentName = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSComponentNameCol).Value
       GEARSDatabase = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSDatabaseCol).Value
       GEARSServer = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSServerCol).Value
       GEARSNSLookup = Worksheets("GEARS").Cells(GEARSRowCnt, GEARSNSLookupCol).Value
       
         If GEARSNSLookup = "Not Found" And GEARSDatabaseCol <> "" Then
           ' ****TODO:Loop through Cyber Worksheet to find new database server where GEARSComponentName = CyberComponentName and CyberComputer contains "DB-" or "-DB"
           ' ****TODO:If found then add a line to ServerSWUpdates
           ' ****TODO:Add GEARSServerSWUpdatesMarkedCol - fill with servername from ServerSWUpdates
       End If
   Next GEARSRowCnt
     
   
   
End Sub


'Sub ReplaceGEARSSameSoftware()
  ' ****TODO: Loop through GEARS ignoring already marked
  ' ****TODO: If the component is NSLookup = "Not found" then
  ' ****TODO: If same software on all servers look for cyber for all servers to replace "Not Found" servers
  ' ****TODO: Add GEARSServerSWUpdatesMarkedCol - fill with servername from ServerSWUpdates
  
'End Sub



