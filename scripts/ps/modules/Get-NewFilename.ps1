function Get-NewFilename
{
    [CmdletBinding(ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true, 
                   HelpMessage = 'Filename to increment')]
        [string]$Path
    )
    if (Test-Path $Path -PathType Leaf) 
    {
        $fileIncrement = 1
        $origPath = $Path
        do {
            $Path = "{0}\{1}_{2}{3}" -f 
                (Split-Path $origPath -Parent),
                [System.IO.Path]::GetFileNameWithoutExtension($origPath),
                $fileIncrement,
                [System.IO.Path]::GetExtension($origPath)
            $fileIncrement = $fileIncrement + 1
        } while (Test-Path $Path -PathType Leaf) 
    }
    return $Path
}