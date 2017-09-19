Import-Module Posh-SSH

Push-Location -Path "$(Split-Path -Parent $MyInvocation.MyCommand.Path)"

$credentials = Get-Credential

$ipAddressesArray = Get-Content .\CiscoSwitchIPAddress.txt

$commands = Get-Content .\CiscoSwitchCommands.txt

foreach ($ipAddress in $ipAddressesArray)
{
    try
    {
        New-SSHSession -ComputerName $ipAddress -Credential $credentials -AcceptKey -ErrorAction Stop
        $continue = $true
    }
    catch 
    {
        try
        {
            Write-Output "Connection failed, try again with different password!" | Tee-Object .\Output.txt -Append
            New-SSHSession -ComputerName $ipAddress -Credential (Get-Credential -Credential "nha") -AcceptKey -ErrorAction Stop
            $continue = $true
        }
        catch
        {
            Write-Output "Connection could not be made to $ipAddress`: $_" | Tee-Object .\Output.txt -Append
            $continue = $false
        }
    }

    if ($continue)
    {    
        $session = New-SSHShellStream -Index 0

        Write-Output ("Executing {0} lines of commands. Please wait...`n" -f $commands.count)

        foreach ($command in $commands)
        {
            $session.WriteLine($command.ToString())
            Start-Sleep -Seconds 1
        }
        
        Write-Output ("Output on {0}" -f (Get-Date)) | Tee-Object .\Output.txt -Append
        $session.Read() | Tee-Object .\Output.txt -Append

        Get-SSHSession | Remove-SSHSession
    }
}