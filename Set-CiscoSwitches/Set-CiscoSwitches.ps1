<#
    Created by Matt Marchese
    Version 2018.06.29
#>

Import-Module Posh-SSH

Function New-Popup
{
    <#
    .Synopsis
    Display a Popup Message
    .Description
    This command uses the Wscript.Shell PopUp method to display a graphical message
    box. You can customize its appearance of icons and buttons. By default the user
    must click a button to dismiss but you can set a timeout value in seconds to
    automatically dismiss the popup.

    The command will write the return value of the clicked button to the pipeline:
    OK     = 1
    Cancel = 2
    Abort  = 3
    Retry  = 4
    Ignore = 5
    Yes    = 6
    No     = 7

    If no button is clicked, the return value is -1.
    .Example
    PS C:\> new-popup -message "The update script has completed" -title "Finished" -time 5

    This will display a popup message using the default OK button and default
    Information icon. The popup will automatically dismiss after 5 seconds.
    .Notes
    Last Updated: April 8, 2013
    Version     : 1.0

    .Inputs
    None
    .Outputs
    integer

    Null   = -1
    OK     = 1
    Cancel = 2
    Abort  = 3
    Retry  = 4
    Ignore = 5
    Yes    = 6
    No     = 7
    #>

    Param (
        [Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a message for the popup")]
        [ValidateNotNullorEmpty()]
        [string]$Message,

        [Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a title for the popup")]
        [ValidateNotNullorEmpty()]
        [string]$Title,

        [Parameter(Position=2,HelpMessage="How many seconds to display? Use 0 require a button click.")]
        [ValidateScript({$_ -ge 0})]
        [int]$Time=0,

        [Parameter(Position=3,HelpMessage="Enter a button group")]
        [ValidateNotNullorEmpty()]
        [ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNo","YesNoCancel","RetryCancel")]
        [string]$Buttons="OK",

        [Parameter(Position=4,HelpMessage="Enter an icon set")]
        [ValidateNotNullorEmpty()]
        [ValidateSet("Stop","Question","Exclamation","Information" )]
        [string]$Icon="Information"
    )

    #convert buttons to their integer equivalents
    Switch ($Buttons)
    {
        "OK"               {$ButtonValue = 0}
        "OKCancel"         {$ButtonValue = 1}
        "AbortRetryIgnore" {$ButtonValue = 2}
        "YesNo"            {$ButtonValue = 4}
        "YesNoCancel"      {$ButtonValue = 3}
        "RetryCancel"      {$ButtonValue = 5}
    }

    #set an integer value for Icon type
    Switch ($Icon)
    {
        "Stop"        {$iconValue = 16}
        "Question"    {$iconValue = 32}
        "Exclamation" {$iconValue = 48}
        "Information" {$iconValue = 64}
    }

    #create the COM Object
    try
    {
        $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
        #Button and icon type values are added together to create an integer value
        $wshell.Popup($Message,$Time,$Title,$ButtonValue+$iconValue)
    }
    catch
    {
        #You should never really run into an exception in normal usage
        Write-Warning "Failed to create Wscript.Shell COM object"
        Write-Warning $_.exception.message
    }
}

function New-XAMLWindow
{
    [CmdletBinding()]
    param
    (
        $InputXML
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    #Read XAML

    $reader = New-Object System.Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }

    #===========================================================================
    # Load XAML Objects In PowerShell
    #===========================================================================

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)
    }

    $WPFRun_Button.Add_Click({$form.Close()})

    $WPFUse_Last.Add_Click({
        $WPFCommands_Textbox.text = (Get-Content .\CiscoSwitchCommands.txt) -Join("`n")
        $WPFIPAddresses_Textbox.text = (Get-Content .\CiscoSwitchIPAddress.txt) -Join("`n")
    })

    $WPFClear_Button.Add_Click({
        $WPFCommands_Textbox.text = $null
        $WPFIPAddresses_Textbox.text = $null
    })

    $inputBox = $Form.ShowDialog() | Out-Null

    return [PSCustomObject]@{
        "Commands"    = $WPFCommands_Textbox.text
        "IPAddresses" = $WPFIPAddresses_Textbox.text
    }
}

$InputXML = @"
<Window x:Name="CiscoSwitch_MainWindow" x:Class="CiscoSwitchesTemplate.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:CiscoSwitchesTemplate"
        mc:Ignorable="d"
        Title="Cisco Switch Update Script" Height="381.187" Width="641.549">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="261*"/>
            <ColumnDefinition Width="56*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Commands" HorizontalAlignment="Left" Margin="28,21,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.849,0.385" Height="30" Width="146" FontWeight="Bold" FontFamily="Consolas" FontSize="18"/>
        <Label Content="IP Addresses" HorizontalAlignment="Left" Margin="384,21,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.849,0.385" Height="26" Width="123" FontFamily="Consolas" FontWeight="Bold" FontSize="16"/>
        <TextBox x:Name="Commands_Textbox" HorizontalAlignment="Left" Height="228" TextWrapping="Wrap" AcceptsReturn="True" VerticalAlignment="Top" Width="331" Margin="28,56,0,0" Cursor="IBeam" TabIndex="0"/>
        <TextBox x:Name="IPAddresses_Textbox" HorizontalAlignment="Left" Height="228" TextWrapping="Wrap" AcceptsReturn="True" VerticalAlignment="Top" Width="214" Margin="384,56,0,0" Cursor="IBeam" Grid.ColumnSpan="2" TabIndex="1"/>
        <Button x:Name="Run_Button" Content="Run" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="1,302,0,0" TabIndex="2" Height="20" Grid.Column="1"/>
        <Button x:Name="Use_Last" Content="Use Last Commands" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="28,302,0,0"/>
        <Button x:Name="Clear_Button" Content="Clear" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="147,302,0,0"/>

    </Grid>
</Window>
"@

Push-Location -Path "$(Split-Path -Parent $MyInvocation.MyCommand.Path)"

$credentials = Get-Credential

$exitScript = $false

do
{
    $InputWindowInformation = New-XAMLWindow -InputXML $InputXML

    $ipAddressesArray = $InputWindowInformation.IPAddresses.Split("`n").Trim()
    $ipAddressesArray | Out-File .\CiscoSwitchIPAddress.txt

    $commands = $InputWindowInformation.Commands.Split("`n").Trim()
    $commands | Out-File .\CiscoSwitchCommands.txt

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
                New-SSHSession -ComputerName $ipAddress -Credential (Get-Credential) -AcceptKey -ErrorAction Stop
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

    $popUpWindow = New-Popup -Message "Would you like to run the script again?" -Title "Cisco Switch Update Script" -Buttons "YesNo" -Icon Question

    if ($popUpWindow -eq 6)
    {
        $exitScript = $false
    }
    elseif ($popUpWindow -eq 7)
    {
        $exitScript = $true
    }

} while ($exitScript -eq $false)
