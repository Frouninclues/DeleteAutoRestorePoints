' VBScript to run a PowerShell script from memory using -EncodedCommand

Set objShell = CreateObject("WScript.Shell")

' PowerShell script content as a multi-line string, preserving exact formatting
ps1Content = "function Remove-ComputerRestorePoint {" & vbCrLf & _
"    [CmdletBinding(SupportsShouldProcess = $True)]" & vbCrLf & _
"    param(" & vbCrLf & _
"        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]" & vbCrLf & _
"        $RestorePoint" & vbCrLf & _
"    )" & vbCrLf & _
"    begin {" & vbCrLf & _
"        $fullName = ""SystemRestore.DeleteRestorePoint""" & vbCrLf & _
"        # Check if the type is already loaded" & vbCrLf & _
"        $isLoaded = $null -ne ([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetTypes() } | Where-Object { $_.FullName -eq $fullName })" & vbCrLf & _
"        if (!$isLoaded) {" & vbCrLf & _
"            Add-Type -MemberDefinition @""" & vbCrLf & _
"[DllImport (""Srclient.dll"")]" & vbCrLf & _
"public static extern int SRRemoveRestorePoint (int index);" & vbCrLf & _
"""@ -Name DeleteRestorePoint -Namespace SystemRestore -PassThru" & vbCrLf & _
"        }" & vbCrLf & _
"    }" & vbCrLf & _
"    process {" & vbCrLf & _
"        foreach ($restorePoint in $RestorePoint) {" & vbCrLf & _
"            if ($PSCmdlet.ShouldProcess(""$($restorePoint.Description)"", ""Deleting Restore Point"")) {" & vbCrLf & _
"                [SystemRestore.DeleteRestorePoint]::SRRemoveRestorePoint($restorePoint.SequenceNumber) | Out-Null" & vbCrLf & _
"            }" & vbCrLf & _
"        }" & vbCrLf & _
"    }" & vbCrLf & _
"}" & vbCrLf & _
"" & vbCrLf & _
"Get-ComputerRestorePoint | Where-Object {$_.RestorePointType -ne 16} | Remove-ComputerRestorePoint"

' Encode the script to Base64 (UTF-16LE without BOM)
encodedScript = Base64Encode(ps1Content)

' Build the PowerShell command with -EncodedCommand
command = "powershell.exe -ExecutionPolicy Bypass -EncodedCommand " & encodedScript

' Run the PowerShell script hidden (window style 0), waiting for completion
objShell.Run command, 0, True

' Clean up objects
Set objShell = Nothing

' Function to encode string to Base64
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

' Function to convert string to UTF-16LE binary, skipping BOM
Function Stream_StringToBinary(Text)
    Const adTypeText = 2
    Const adTypeBinary = 1
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    BinaryStream.Type = adTypeText
    BinaryStream.Charset = "utf-16le"
    BinaryStream.Open
    BinaryStream.WriteText Text
    BinaryStream.Position = 0
    BinaryStream.Type = adTypeBinary
    BinaryStream.Position = 2  ' Skip the BOM (2 bytes)
    Stream_StringToBinary = BinaryStream.Read
    Set BinaryStream = Nothing
End Function