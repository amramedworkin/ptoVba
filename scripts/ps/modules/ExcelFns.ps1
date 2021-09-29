function Convert-ExcelFormat
{
    param(
        [Parameter(Mandatory = $true, 
                   HelpMessage = 'Source file (usually XLST')]
        [string]$Source,
        [Parameter(Mandatory = $true, 
                   HelpMessage = 'Target format (XLSM supported)')]
        [string]$Format,
        [Parameter(Mandatory = $true, 
                  HelpMessage = 'File name to change (if not specified, source name)')]
        [string]$Output,
        [Parameter(HelpMessage = 'Rename the first sheet')]
        [string]$SheetName = $null,
        [Parameter(Mandatory = $true, 
                  HelpMessage = 'Delete output file if exists')]
        [switch]$Force

    )
    # Need literal path for rest of this 
    if ( !(Test-Path -Path $Source -PathType Leaf) )
    {
        Write-Error "$Source does not exist.  Process aborted"
        return $null
    }

    $xlFormat = $null
    switch ($Format) 
    {
        xlsm { $xlFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled }
        defalt { $xlFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault }
    }

    $excel = New-Object -ComObject Excel.Application
    try 
    {
        if (Test-Path $Output )
        {
            if ($Force) { Remove-Item -Path $Output -Force }
            else { $Output = Get-NewFilename }
        }
	    $book = $excel.Workbooks.Open($Source)
        # Rename 1st sheet if required
        if ($SheetName) { $book.worksheets.item(1).Name = $SheetName }
        $book.SaveAs($Output,$xlFormat)
        $book.Close()
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

function Close-ExcelWorkbooks
{
	try {
		# attach to running Excel instance
		$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
	} catch {
		if ($_.Exception.HResult -eq -2146233087) {
			# Excel not running (nothing to do)
			return 0
		} else {
			# unexpected error: log and terminate
			Write-EventLog -LogName Application `
			-Source 'Application Error' `
			-EventId 500 `
			-EntryType Error `
			-Message $_.Exception.Message
			return 1
		}
	}

	$xl.DisplayAlerts = $false  # prevent Save() method from asking for confirmation
								# about overwriting existing files (when saving new
								# workbooks)
								# WARNING: this might cause data loss!

	foreach ($wb in $xl.Workbooks) {
		$wb.Save()
		$wb.Close($false)
	}

	$xl.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}