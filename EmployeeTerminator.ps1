## Code to hide the powershell command window when GUI is running
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

## Set up window XML
$inputXML = @"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="MRG Employee Terminator" Height="350" Width="525">
    <Grid>
        <Button x:Name="button" Content="Continue" HorizontalAlignment="Left" Height="22" Margin="407,288,0,0" VerticalAlignment="Top" Width="100"/>
        <Button x:Name="button1" Content="Cancel" HorizontalAlignment="Left" Height="22" Margin="300,288,0,0" VerticalAlignment="Top" Width="102"/>
        <TextBlock x:Name="textBlock" HorizontalAlignment="Left" Height="23" Margin="0,10,0,0" TextWrapping="Wrap" Text="Authorized User Name:" VerticalAlignment="Top" Width="127"/>
        <TextBlock x:Name="textBlock1" HorizontalAlignment="Left" Height="22" Margin="0,33,0,0" TextWrapping="Wrap" Text="Password:" VerticalAlignment="Top" Width="127"/>
        <TextBlock x:Name="textBlock2" HorizontalAlignment="Left" Height="23" Margin="0,104,0,0" TextWrapping="Wrap" Text="Employee to be terminated" VerticalAlignment="Top" Width="146"/>
        <TextBlock x:Name="textBlock3" HorizontalAlignment="Left" Height="21" Margin="0,130,0,0" TextWrapping="Wrap" Text="User Name:" VerticalAlignment="Top" Width="63"/>
        <TextBox x:Name="textBox" HorizontalAlignment="Left" Height="23" Margin="132,10,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="176"/>
        <PasswordBox x:Name="passwordBox" HorizontalAlignment="Left" Height="22" Margin="132,33,0,0" VerticalAlignment="Top" Width="176"/>
        <TextBox x:Name="textBox1" HorizontalAlignment="Left" Height="23" Margin="68,127,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
        <Image x:Name="image" HorizontalAlignment="Left" Height="133" Margin="325,104,0,0" VerticalAlignment="Top" Width="163" IsEnabled="True" Source="C:\Users\brismith\Desktop\WarningImage.png" Visibility="Hidden"/>
        <Button x:Name="button2" Content="Enable Changes" HorizontalAlignment="Left" Height="23" Margin="68,167,0,0" VerticalAlignment="Top" Width="91"/>
        <TextBlock x:Name="textBlock4" HorizontalAlignment="Left" Height="69" Margin="10,219,0,0" TextWrapping="Wrap" Text="WARN: Changes are ENABLED pressing Execute will terminate this employees access from MRG controled IT Resources" VerticalAlignment="Top" Width="293" Foreground="Red" Visibility="Hidden"/>

    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
#Get-FormVariables
 
#===========================================================================
# Actually make the objects work
#===========================================================================

##Setup Global Support Materials & Functions
$wshell = New-Object -ComObject Wscript.Shell ##Create a Windows Scripting host shell instance to support interactive popups.
function SanityCheckMod{
    if ($WPFimage.Visibility -ne 'Visible'){
        $WPFimage.Visibility = "Visible"
        $WPFTextBlock4.Visibility = "Visible"
    }
    else{
        $WPFimage.Visibility = "Hidden"
        $WPFTextBlock4.Visibility = "Hidden"
    }
}
function CheckSanity {
    if($WPFimage.Visibility -eq 'Visible'){
        return $true
    }
    else{
        return $false
    }

}
function RemoveUser{
    #Generate Secured Credental object from user provided resoruces.
    $AuthorizedPassword = $WPFpasswordBox.Password
    $AuthorizedUserName = $WPFtextBox.Text

    $SecuredPassword = ConvertTo-SecureString $AuthorizedPassword -AsPlainText -Force
    $ActionCredentals = New-Object System.Management.Automation.PSCredential($AuthorizedUserName, $SecuredPassword)
    
    #Change Users AD Account password to randomstring to lock user out off office 365 with forced sync
    $PasswordcharPool = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789~!@#$%^&*()_+=-~".ToCharArray()
    $NewRandomPassword = ""
    for ($i = 0; $i -lt 54; $i++){
        $NewRandomPassword += $PasswordcharPool | Get-Random
    }
    
    $RandomPassword = ConvertTo-SecureString $NewRandomPassword -AsPlainText -Force
    $TerminatedUserName = $WPFtextbox1.Text

    Set-ADAccountPassword -Identity $TerminatedUserName -Credential $ActionCredentals -Reset -NewPassword $RandomPassword -PassThru
    #Disable Users AD Account
    Disable-ADAccount -Identity $TerminatedUserName -Credential $ActionCredentals -PassThru
    
    #create and resolve spiceworks ticket with 1hr worked
    $mycredentials = Get-Credential -Message "Please provide the email address and password accociated with your ticketing purposes"
    $SmtpServer = 'smtp.office365.com'
    $MailtTo = 'helpdesk@medfordradiology.com'
    $MailFrom = $mycredentials.UserName 
    $MailSubject = "Disabled User: " + $WPFtextbox1.Text
    $emailbody = "Disabled user account for: " + $WPFtextbox1.Textt + " Created. Please contact IT for more infomraiton `n #add 1h `n #close"

    Send-MailMessage -To "$MailtTo" -from "$MailFrom" -Subject $MailSubject -Body $emailbody -SmtpServer $SmtpServer -UseSsl -Port 587 -Credential $mycredentials 
    
    #add code to notify HR & contact partnets about change in employment.

}
#Configure Cancel Button
$WPFbutton1.Add_Click({$form.Close()})

#Configure User Sanity Check Button
$WPFbutton2.Add_Click({SanityCheckMod})

#Configure Continue Button
$WPFbutton.Add_Click({
    if(CheckSanity){
        #The Remove User command below is disabled for testing
        RemoveUser
        $wshell.Popup("SUCCESS.", 5, "MRG Employee Terminator", 0x30)
    }
    else{
        $wshell.Popup("Changes are not enabled, No changes Occured.", 5, "MRG Employee Terminator: Change Enabled ERROR", 0x30)
    }
})


#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog()