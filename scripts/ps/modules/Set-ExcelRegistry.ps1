function Set-ExcelRegistry
{
    Write-Output("------------------------------------------------")
    Write-Output("REGEDIT SETTING Excel Security Registry Values")
    try
    {
        $excel = New-Object -ComObject Excel.Application
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
    }
    catch {
        Write-Error $_
    }
    finally 
    {
        Write-Output("REGEDIT COMPLETE")
        $excel.Quit()
        # http://technet.microsoft.com/en-us/library/ff730962.aspx
        [Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
