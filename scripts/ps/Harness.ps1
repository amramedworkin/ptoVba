#region HK
#region hk.MODULES
[string]$modulePath = $PSScriptRoot + "\modules"
write-host "================================"
write-host "================================"
write-host "================================"
write-host "================================"
write-host "================================"
write-host "================================"
write-host "================================"
# Install the other script files
# - Eventually convert to a module
(Get-ChildItem -Path $modulePath -Filter *.ps1) | ForEach-Object { . $_.Fullname }
#endregion hk.MODULES

# HK
<# ACTIONS
    export = Export vba objects to a folder for later import
    export.force = Delete vba objects already in folder
    import = Import vba from folder (export) into an book...save as XLSM
    regedit = Changes to registry settings to make this work
#>

#-------------------------
#-------------------------
#-------------------------
<#
Step 1 - download latest ecmo report	tbd
Step 2 - Migrate VBA code into workbook
Step 3 - run PrepWorkbook vba code	obsolete
Step 4 - add gears data from arch server sw, arch database, archserverswunknown and archserverdatabaseunknown - use AmyBF and Big Fix views	from spreadsheet
Step 5 - Run an NSLookup powershell script on GEARS data	amy said not necessary
Step 6 - remove .uspto.gov from ecmo and gears tabs	captured in the keys of the servers JSON object
Step 7 - run ExactCybertoGEARSServerMatch vba code	performed by servers object load process
Step 8 - run IdentifyType vba code	performed by section of code marked by IdentifyMatch
Step 9 - Run MoveCyberNonProd vba code	performed by section of code marked by MoveCyberNoProd
Step 10 - Move in SDAP data	from spreadsheet
Step 11 - run MatchOnSDAP code	perform servers object load process
Step 12 - Move in manual tab�	from spreadsheet
Step 13 - RunCheckManual code	perform servers object load process
Step 14 - Download Diamond (formally mysae) data - remove .uspto.gov� �https://diamond.uspto.gov/login�select Hosts select Data Center column layout	from spreadsheet
Step 15 - Run DiamondCheck vba code	perform servers object load process
Step 16 - Download PML (Components List)	
Step 17 - Run FindComponentinPML vba code (takes approx 35 mins)	
Step 18 - download servers list remove .uspto.gov - use BigFix view	
Step 19 - Run ServerSWUpdate vba code	
Step 20 - Run FindServersNSLookup vba code
#>
$actions = @(1,2)
#-------------------------
#-------------------------
#-------------------------

$scriptFolder = Split-Path $MyInvocation.MyCommand.Path -Parent
$vbaFolder =  Resolve-Path ("{0}\{1}" -f $scriptFolder,"\..\vba")
$dataSourceFolder = Resolve-Path ("{0}\{1}" -f $scriptFolder,"..\..\data\test\input")
$vbaSourceFile =  "{0}\{1}" -f $dataSourceFolder,"Reference3.xlsm"
$excelSourceDataFile = "{0}\{1}" -f $dataSourceFolder,"NewEcmoReport.xlsx"
$excelTargetFolder = Resolve-Path ("{0}\{1}" -f $scriptFolder,"..\..\data\test\output")
$excelTargetFile = "{0}\{1}" -f $excelTargetFolder,"Cyber-vba.xlsm"
# Condition the system if Registry not right for Excel object extract

# Only test runs when test is specified
if ($actions.Contains('test')) 
{
    $excel = New-Object -ComObject Excel.Application
    $book = $excel.Workbooks.Open($excelTargetFile)
    Write-Host "Testing $book.Name"
    $excel.Visible=$true
    $excel.DisplayAlerts = $true
    $excel.ScreenUpdating = $true
    Close-ExcelWOrkbooks
    exit
}
if ($actions.Contains('regedit')) { Set-ExcelRegistry }
#endregion HK

#region WORKFLOW
<#
    Step 1 - download latest ecmo report
        - Copy from ecmo.xlst
#>
# STEP 1 - Copy to Cyber.xlsm
if ($Actions.Contains(1)) {
    Write-Host "# STEP 1 - Download (copy) latest Ecmo Report - Copy from ecmo.xlst to Cyber.xlsm"
    # Kill any processes locking the books
    Close-ExcelWOrkbooks
    Convert-ExcelFormat -Source $excelSourceDataFile -Format 'XLSM' -Output $excelTargetFile -SheetName "Cyber" -Force 
}

# Step 2 - Migrate VBA code into workbook
if ($Actions.Contains(2)) {
    Write-Host "# STEP 2.a - Extract VBA"
    # Make sure the output folder exists and is empty
    Remove-Item -Path ("{0}\*" -f $VbaFolder) -Include * -Force -Confirm:$false
    Close-ExcelWOrkbooks
    Export-ExcelProject -WorkbookPath $vbaSourceFile -Output $VbaFolder -IncludeAutoNamed
    Write-Host "# STEP 2.b - Import VBA Code Into Cyber"
    Close-ExcelWOrkbooks
    Import-ExcelProject -vbaFolder $VbaFolder -Path $excelTargetFile
}

# After we've conditioned Cyber.xlsm we keep it open
# Close-ExcelWOrkbooks
$excel = New-Object -ComObject Excel.Application
try 
{
    $book = $excel.Workbooks.Open($excelTargetFile)
    $excel.Visible=$false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $cyberSheet = $book.Worksheets.Item('Cyber')
    
    # Step 3 - run PrepWorkbook vba code
    if ($Actions.Contains(3)) 
    {
        $excel.Run("PrepWorkbook") 
    }
    if ($Actions.Contains(5)) 
    {
        Write-Host("STEP 2 - Condition ECMO sheet (uppercase,remove .uspto.gov, rename sheet")
        $computerNameCol = 0
        $computerNameHeader = 'Computer Name'
        for($i = 1; $i -lt 40; $i++) 
        { 
            if ($cyberSheet.Cells.Item(1,$i) -eq $computerNameHeader) 
            {
                $computerNameCol = $i
                break
            }
        }
        if ($computerNameCol -eq 0) 
        { 
            Write-Error "Unable to find $computerNameHeader column in [Cyber]"
            exit 1
        }
        $cyberRange = $cyberSheet.UsedRange
        $cyberRowCount = $cyberRange.Rows.Count
        for($i = 2; $i -lt $cyberRowCount; $i++) 
        {
            $cyberSheet.Cells.Item($i,$computerNameCol).Value2 = cyberSheet.Cells.Item($i,$computerNameCol).Value2 -replace '.uspto.gov',''
        }
        $book.Save()
    }
}
catch 
{
    Write-Error $_
}
finally {
    $excel.Quit()
    # http://technet.microsoft.com/en-us/library/ff730962.aspx
    [Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
#endregion WORKFLOW



Write-Output("====================================================")