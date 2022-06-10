<#
    .Synopsis
    Editor for Citrix ADC LDAP One Time Password Secret
    .DESCRIPTION
    Editor for Citrix ADC LDAP One Time Password Secret
    .Note
    New Format with Citrix ADC 13.41!
    https://docs.citrix.com/en-us/citrix-adc/13/aaa-tm/native-otp-authentication/otp-encryption-tool.html
#>

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration
Add-Type -AssemblyName PresentationFramework

Import-Module "$PSScriptRoot\inc\NITWPFTools.psm1" -Force
Import-Module "$PSScriptRoot\inc\New-WPFJobWindow.psm1" -Force
Import-Module "$PSScriptRoot\inc\New-XMLSettingsDialog.psm1" -Force
Import-Module "$PSScriptRoot\inc\OTPQRCodes.psm1" -Force
Import-Module "$PSScriptRoot\inc\New-OtpQRCodeWindow.psm1" -Force
Import-Module "$PSScriptRoot\Send-QRCodeEmail.ps1" -Force


#Load Config
$Global:Configxml = $Null 
#Test 4 Userconfig
$OTPConfigFile = $env:APPDATA + "\AndreasNick\OTPEdit\OTPEditConfig.xml"


function Set-GlobalSettings{

  $Global:AttributeStore = $Configxml.OtpEditConfig.AttributeStore
  $Global:ODTDefinitionString = $Configxml.OtpEditConfig.ODTDefinitionString 
  $Global:ODTSeperator = $Configxml.OtpEditConfig.ODTSeperator
  $Global:LDAPServer = $Configxml.OtpEditConfig.LDAPServer 
}

if((Get-Module -ListAvailable -Name ActiveDirectory) -eq $null){
  
  Out-Message -Message "Can't start - You need the Active Directory PowerShell Module!`n Please edit the config file" -Type Error -ParentWindow $null
  #Add-Type -AssemblyName System.Windows.Forms
  #[System.Windows.Forms.MessageBox]::Show("Can't start - You need the Active Directory PowerShell Module!`n Please edit the config file","Information",0,64)
  break
}



  $Global:ConfigXml = New-Object xml
  
  if(-not (Test-Path $OTPConfigFile)) {
    if(-not (Test-Path (Split-Path $OTPConfigFile -Parent))){
      New-Item -Path (Split-Path $OTPConfigFile -Parent) -ItemType Directory -Force
    }
    #LoadMainConfig
    $Global:Configxml.Load($($PSScriptRoot + "\OTPEditConfig.xml"))
    $Global:Configxml = New-XMLSettingsDialog -Title "Settings" -SettingsXml $Configxml -EntryKey "OtpEditConfig" -ParentWindow $window
    $Global:Configxml.Save($OTPConfigFile)
  }  else 
  {
    $Global:ConfigXml.Load($OTPConfigFile)
  }
  
  Set-GlobalSettings













#
#
# Add new Device
#
#

function AddDeviceToUser($SamAccountName){

import-module activedirectory

$LocalSite = (Get-ADDomainController -Discover).Site
$NewTargetGC = Get-ADDomainController -Discover -Service 6 -SiteName $LocalSite
IF (!$NewTargetGC)
{ $NewTargetGC = Get-ADDomainController -Discover -Service 6 -NextClosestSite }
$NewTargetGCHostName = $NewTargetGC.HostName
$LocalGC = “$NewTargetGCHostName” + “:3268”
 
$User = Get-ADUser -Server $LocalGC -filter { sAMAccountName -eq $SamAccountName }  -Properties *
  
  
  $EMailAdress = $null
  
  if($User.Mail -eq $null){
    if($User.UserPrincipalName -eq $null){
      $EMailAdress = 'unknown@unknown.local'
    } else {
      $EMailAdress = $User.UserPrincipalName
    }
  } else {
    $EMailAdress = $User.Mail
  }
  
  $Username = $null
  if($user.UserPrincipalName -eq ""){
    $userName = $user.SamAccountName
  } else 
  {
    $username = $user.UserPrincipalName
  }
  
  $pass = Get-HexPass
  $Base32Pass = Convert-ToHexToBase32 $pass

  $result = New-OtpQRCodeWindow -Secret $Base32Pass -Settings $Global:Configxml -Parent $window -Email $EMailAdress -UserPrincipalName $username -Device $Global:Configxml.OtpEditConfig.OTPDefaultDeviceName
  
  if($Result -ne $null)
  {
  
    $NewDevString = $( $Result.Device + "=" + $Result.secret + '&,') #NewDevice
    Write-output "Add Device $DeviceString"
  
   
    $DeviceString = $null
    $DeviceString = $user.item($AttributeStore)
    $DeviceStringLength =$user.item($AttributeStore).length
    if($DeviceStringLength -lt 1){$DeviceString = $null}
    Write-Output $DeviceString

    
    #Device Exist?
    if($DeviceString -notmatch $($Result.Device)){
    
      if($DeviceString -notmatch $('^' + $Global:Configxml.OtpEditConfig.ODTDefinitionString)){
        $DeviceString = $($Global:Configxml.
        OtpEditConfig.ODTDefinitionString + $DeviceString)
      }
    
      $Command = 'Set-ADUser -Identity "' + $User.DistinguishedName + '" -Replace @{'''+ $AttributeStore +'''="'+$DeviceString+$NewDevString+'"}'
  
      #$res = Out-Message -Message $("SamAccountName = $User `nExecute command : " + $Command) -Type Information -ParentWindow $window -DisableCancel $false
      $res="OK"
  
      if($res -eq "OK"){
        try{
          write-Output $Command
          Invoke-Expression  $Command -ErrorAction Stop
       
          #
          # Send QRCode
          #

          #$res = Out-Message -Message $("Send E-Mail to User $User (" + $Result.EMail +")") -Type Information -ParentWindow $window -DisableCancel $false
          $res="OK"

          if($res -eq "OK"){
          
            [int] $Port = $($Global:Configxml.OtpEditConfig.SMTPPort.InnerText)

            Send-QRCodeEMail -SMTPPort $Port  -SMPTServer $($Global:Configxml.OtpEditConfig.SMTPServer.InnerText) `
            -From $($Global:Configxml.OtpEditConfig.SMTPMailFrom.InnerText) `
            -Subject $($Global:Configxml.OtpEditConfig.SMTPSubject.InnerText) -SMTPUseSSL ($Global:Configxml.OtpEditConfig.SMTPUseSSL.InnerText -eq "true") `
            -SMTPMailuser $($Global:Configxml.OtpEditConfig.SMTPUser.InnerText) -SMTPMailPassword $($Global:Configxml.OtpEditConfig.SMTPPassword.InnerText) `
            -To $Result.EMail -QRCode $result.QRCode -UserName $User -Secret $Result.Secret -DeviceName $Result.Device -UserPrincipleName $Result.UserPrincipleName
          
            <#
            
                $ImageBase64 = [Convert]::ToBase64String($result.QRCode)

                $Mailuser =  $Global:Configxml.OtpEditConfig.SMTPUser.InnerText
                $Mailpwd = $Global:Configxml.OtpEditConfig.SMTPPassword.InnerText
                $MailPort = $Global:Configxml.OtpEditConfig.SMTPPort.InnerText
            
                $secure_pwd = $Mailpwd  | ConvertTo-SecureString -AsPlainText -Force

                $creds = New-Object System.Management.Automation.PSCredential -ArgumentList  $Mailuser, $secure_pwd

                #
                # Als als html Mail
                #
                [String] $htmlDoc = $null
                #Bachground

                $htmlDoc += '<style>'
                $htmlDoc += 'body {background-color:#d2E0EF;}'
                $htmlDoc += 'h1   {color: blue;}'
                #$htmlDoc += 'img {width: 50%;height: auto;}'
                $htmlDoc += 'strong    {color:blue;}'
                $htmlDoc += '* {font-family: Consolas;}'
                $htmlDoc += '</style>'
                $htmlDoc += '</head>'
                $htmlDoc += '<body>'

                #Headline
                $htmlDoc += '<h1>OTP QR Code for: ' + $User + '</h1>'
                $htmlDoc += '<h2>Please scan the code with an Authenticator (Microsoft, Google etc.)</h2>'

                $htmlDoc += '<img src="data:image/png;base64,'+ $ImageBase64 +'" />'
                $htmlDoc += '</body>'
                #$htmlDoc | out-file C:\Profiles\Administrator\Desktop\test2.html

                Send-MailMessage -From $($Global:Configxml.OtpEditConfig.SMTPMailFrom.InnerText)  -To 'Andreas <a.nick@nick-it.de>' -Subject  $($Global:Configxml.OtpEditConfig.SMTPSubject.InnerText)  `
                -SmtpServer  $($Global:Configxml.OtpEditConfig.SMTPServer.InnerText) -Port  $MailPort -Credential $creds `
                -Body $htmlDoc -UseSsl:($Global:Configxml.OtpEditConfig.SMTPUseSSL.InnerText -eq "true")  -Encoding UTF8 -BodyAsHtml -ErrorAction Stop

            #>
          }
        
        } catch {
          Out-Message -Message $($_ | Out-String) -Type Error -ParentWindow $window
        }
      } 
      
    } else
    {
      Out-Message -Message $("I cannot create the device, because a device with the same name already exists :"+ $Result.Device) -Type Error -ParentWindow $window
    }
  
    #Write-Host $($result.Secret)
  }
  
  
}

$paramuser=$args[0]
If ($paramuser -eq $null) {
	$paramuser = read-host -prompt "Benutzername (z.B. dozent)"
}

AddDeviceToUser $paramuser


