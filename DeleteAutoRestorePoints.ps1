function Remove-ComputerRestorePoint {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        $RestorePoint
    )
    begin {
        $fullName = "SystemRestore.DeleteRestorePoint"
        # Check if the type is already loaded
        $isLoaded = $null -ne ([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.FullName -eq $fullName })
        if (!$isLoaded) {
            Add-Type -MemberDefinition @"
[DllImport ("Srclient.dll")]
public static extern int SRRemoveRestorePoint (int index);
"@ -Name DeleteRestorePoint -Namespace SystemRestore -PassThru
        }
    }
    process {
        foreach ($restorePoint in $RestorePoint) {
            if ($PSCmdlet.ShouldProcess("$($restorePoint.Description)", "Deleting Restore Point")) {
                [SystemRestore.DeleteRestorePoint]::SRRemoveRestorePoint($restorePoint.SequenceNumber) | Out-Null
            }
        }
    }
}

Get-ComputerRestorePoint | Where-Object {$_.RestorePointType -ne 16} | Remove-ComputerRestorePoint