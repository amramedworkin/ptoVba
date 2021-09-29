Add-Type -AssemblyName Microsoft.Office.Interop.Excel
# Import programmable objects from a file folder into an Excel spreadsheet...output as XLSM
function Import-ExcelProject
{
    [CmdletBinding(ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true,
                   HelpMessage = 'Folder containing the VBA code to copy in')]
        [string]$VbaFolder,
        [Parameter(Mandatory = $true,
                   HelpMessage = 'Source of data XLST (or other) file')]
        [string]$Path
    )
	$excel = New-Object -ComObject Excel.Application
    try 
    {
	    $book = $excel.Workbooks.Open($Path)
	    $excel.Visible=$true
	    $excel.DisplayAlerts = $false
	    $excel.ScreenUpdating = $false
        $vbaFiles = (Get-ChildItem -Path "$VbaFolder\*" -Include *.bas)
        
	    $vbaFiles | ForEach-Object {
			$moduleFilename = $_.FullName
            $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($moduleFilename)
            Write-Host('Importing {0}' -f $moduleFilename)
            if ($book.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }) 
            {
                $book.VBProject.VBComponents.Remove($moduleName)
            }
            # $xlmodule = $book.VBProject.VBComponents.Add(1)
            # $xlmodule.Name = $moduleName
            $book.VBProject.VBComponents.Import($moduleFilename)
		}
		
        # Time to do the heavy lift
        $book.Save()

	    Write-Verbose("Closing $Path")
	    $excel.Workbooks.Close()
    } 
    catch 
    {
        Write-Error $_
    }

    finally
    {
        $excel.Quit()
        # http://technet.microsoft.com/en-us/library/ff730962.aspx
        [Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
