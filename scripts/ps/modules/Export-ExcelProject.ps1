function Export-ExcelProject
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true, 
                   HelpMessage = 'Specifies the path to the Excel Workbook file')]
        [string]$WorkbookPath,
        [Parameter(HelpMessage = 'Specifies export directory')]
        [string]$OutputPath,
        [Parameter(HelpMessage = 'Regular expression pattern identifying modules to be excluded')]
        [string]$Exclude,
        [Parameter(HelpMessage = 'Export items that may be auto-named, like Class1, Module2, etc.')]
        [switch]$IncludeAutoNamed = $false,
        [switch]$Force = $false
    )
    
    function Get-MD5Hash($filePath)
    { 
        $bytes = [IO.File]::ReadAllBytes($filePath)
        $hash = [Security.Cryptography.MD5]::Create().ComputeHash($bytes)
        [BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()
    }
    
    $mo = Get-ItemProperty -Path HKCU:Software\Microsoft\Office\*\Excel\Security `
                           -Name AccessVBOM `
                           -EA SilentlyContinue | `
              Where-Object { !($_.AccessVBOM -eq 0) } | `
              Measure-Object

    if ($mo.Count -eq 0)
    {
        Write-Warning 'Access to VBA project model may be denied due to security configuration.'
    }

    Write-Verbose 'Starting Excel'
    $xl = New-Object -ComObject Excel.Application -EA Stop
    Write-Verbose "Excel $($xl.Version) started"
    $xl.DisplayAlerts = $false
    $missing = [Type]::Missing
    $extByComponentType =  @{ 100 = '.cls'; 1 = '.bas'; 2 = '.cls' }
    $outputPath = ($outputPath, (Get-Item .).FullName)[[String]::IsNullOrEmpty($outputPath)]
    mkdir -EA Stop -Force $outputPath | Out-Null
    
    try
    {
        # Open(Filename, [UpdateLinks], [ReadOnly], [Format], [Password], [WriteResPassword], [IgnoreReadOnlyRecommended], [Origin], [Delimiter], [Editable], [Notify], [Converter], [AddToMru], [Local], [CorruptLoad]) 
        $wb = $xl.Workbooks.Open($workbookPath)
        # foreach($sheet in $wb.worksheets) {
        #     $x = $sheet
        #     $y = 1
        # }
<#        , $false, $true, `
                                 $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, `
                                 $true)
#>
        
        $wb | Get-Member | Out-Null # HACK! Don't know why but next line doesn't work without this
        
        $project = $wb.VBProject
        
        if ($null -eq $project)
        {
            Write-Verbose 'No VBA project found in workbook'
        }
        else
        {
            $tempFilePath = [IO.Path]::GetTempFileName()

            $vbcomps = $project.VBComponents
            
            if (![String]::IsNullOrEmpty($exclude))
            {
                $verbose = ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose') -and $PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
                if ($verbose) 
                {
                    $vbcomps | Where-Object { $_.Name -match $exclude } | ForEach-Object { Write-Verbose "$($_.Name) will be excluded" }
                }
                $vbcomps = $vbcomps | Where-Object { $_.Name -notmatch $exclude }
            }
            
            $vbcomps | ForEach-Object `
            { 
                $vbcomp = $_
                $name = $vbcomp.Name
                $ext = $extByComponentType[$vbcomp.Type]
                if ($null -eq $ext)
                {
                    Write-Verbose "Skipped component: $($name)"
                }
                elseif (!$includeAutoNamed -and $name -match '^(Form|Module|Class|Sheet)[0-9]+$')
                {
                    Write-Verbose "Skipped possibly auto-named component: $name"
                }
                elseif ($vbcomp.CodeModule.CountOfLines -eq 0) {
                    Write-Verbose "Skipped zero code module: $name"
                }
                else
                {
                    $vbcomp.Export($tempFilePath)
                    
                    $exportedFilePath = Join-Path $outputPath "$name$ext"
                    $exists = Test-Path $exportedFilePath -PathType Leaf
                    
                    if ($exists) 
                    { 
                        $oldHash = Get-MD5Hash $exportedFilePath 
                        $newHash = Get-MD5Hash $tempFilePath
                        $changed = !($oldHash -eq $newHash)
                        $status  = ('Unchanged', 'Conflict', 'Unchanged', 'Changed')[[int]$changed + (2 * [int]$force.IsPresent)]
                    }
                    else
                    {
                        $status = 'New'
                    }

                    if (($status -eq 'Changed' -or $status -eq 'New') `
                        -and $pscmdlet.ShouldProcess($name))
                    {
                        Move-Item -Force $tempFilePath $exportedFilePath
                    }
                    
                    New-Object PSObject -Property @{
                        Name   = $name;
                        Status = $status;
                        File   = (Get-Item $exportedFilePath -EA Stop);
                    }
                }
            }        
        }
        $wb.Close($false, $missing, $missing)
    }
    catch 
    {
        Write-Error $_
    }

    finally
    {    
        $xl.Quit()
        # http://technet.microsoft.com/en-us/library/ff730962.aspx
        [Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$xl) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
