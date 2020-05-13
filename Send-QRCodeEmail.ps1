Function Send-QRCodeEMail 
{
  <#
      .SYNOPSIS
       A simple function to send notifications with a QRCode
  #>

  [CmdletBinding()]
  param
  (
    
    [Parameter(Mandatory=$true)] [String] $SMPTServer,
    [Parameter(Mandatory=$true)] [int] $SMTPPort,
    [bool] $SMTPUseSSL = $True,
    [String] $SMTPMailuser,
    [String] $SMTPMailPassword,
    [String] $From = 'administrator@unternehmen.local',
    [Parameter(Mandatory=$true)] [String] $To,
    [String] $Subject = "QRCode Mail",
    [String] $UserName = "Rudi",
    [String] $Secret,
    [String] $DeviceName,
    [String] $UserPrincipleName,
    [Parameter(Mandatory=$true)] [Byte[]] $QRCode
    

  )
  $ImageBase64 = [Convert]::ToBase64String($QRCode)
  $secure_pwd = $SMTPMailPassword  | ConvertTo-SecureString -AsPlainText -Force
  $creds = New-Object System.Management.Automation.PSCredential -ArgumentList   $SMTPMailuser, $secure_pwd

  #Create URL
  $Auth = 'otpauth://totp/'+[System.Web.HttpUtility]::UrlEncode($UserPrincipleName)+'?secret=' + $secret + '&device=' + $DeviceName
  
  #write-Host $Auth 

  #
  # Als als html Mail
  #
  [String] $htmlDoc = $null
  #Bachground

  $htmlDoc += '<style>'
  $htmlDoc += 'body {background-color:#d2E0EF;}'
  $htmlDoc += 'h1   {color: blue;}'
  $htmlDoc += 'strong    {color:blue;}'
  $htmlDoc += '* {font-family: Consolas;}'
  $htmlDoc += '</style>'
  $htmlDoc += '</head>'
  $htmlDoc += '<body>'
  #Headline
  $htmlDoc += '<h1>OTP QR Code for: ' + $UserName + '</h1>'
  $htmlDoc += '<h2>Please scan the code with an Authenticator (Microsoft, Google etc.)</h2>'
  $htmlDoc += '<img src="data:image/png;base64,'+ $ImageBase64 +'" />'
  $htmlDoc += '<p><p>Secret:' + $Secret
  $htmlDoc += '<p><p><a href="' + $Auth + '">Auth Link</a>'

  $htmlDoc += '</body>'

  
      Send-MailMessage -From $From  -To $To -Subject  $Subject -SmtpServer  $SMPTServer -Port  $SMTPPort -Credential $creds `
                   -Body $htmlDoc -UseSsl:$SMTPUseSSL  -Encoding UTF8 -BodyAsHtml -ErrorAction Stop
                   

}