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



$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   MinHeight="350"
   MinWidth="650"
   Height="50"
   ResizeMode="CanResizeWithGrip"
    Width="730"
   Name="OtpWindow"
   Title="OTP-Explorer"
   Topmost="False">
   
    <Grid Margin="10,10,10,10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="300" />
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="5"/>
            

        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"></RowDefinition>
            <RowDefinition Height="30"></RowDefinition>
            <RowDefinition Height="*"></RowDefinition>
            <RowDefinition Height="Auto"></RowDefinition>
        </Grid.RowDefinitions>
        <Grid  Grid.Column="0" Grid.Row="0" Grid.ColumnSpan="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"  />
                <ColumnDefinition Width="120"/>
            </Grid.ColumnDefinitions>

            <ToolBar ToolBarTray.IsLocked="True"  DockPanel.Dock="Top" VerticalAlignment="Top" HorizontalAlignment="Stretch" Grid.Column="0"  >

                <Button  Margin="10,0,0,0" Name="ButtonRefresh">
                    <StackPanel Orientation="Horizontal">
                        <Image  Name="ImageRefresh" Width="32" Height="32" />
                        <TextBlock VerticalAlignment="Center" Margin="4,0,0,0">Refresh</TextBlock>

                    </StackPanel>
                </Button>
                <TextBlock VerticalAlignment="Center" Margin="10,0,0,0">Filter:</TextBlock>
                <TextBox Name="TextBoxADFilter" VerticalAlignment="Center"  Width="200" Margin="10,0,2,2">*</TextBox>
                <Button VerticalAlignment="Center" Margin="10,0,0,0" Name="ButtonExportCSV">ExportCSV</Button>
                <Button  Margin="10,0,0,0" Name="ButtonConfig">
                    <StackPanel Orientation="Horizontal">
                        <Image  Name="ImageConfig" Width="32" Height="32" />
                        <TextBlock VerticalAlignment="Center" Margin="4,0,0,0">Config</TextBlock>
                    </StackPanel>
                </Button>

            </ToolBar>

            <ToolBar ToolBarTray.IsLocked="True"   DockPanel.Dock="Top" VerticalAlignment="Top" Grid.Column="1" >
                 <Button VerticalAlignment="Center"   Margin="10,0,0,0" Name="ButtonAbout">
                    <StackPanel Orientation="Horizontal">
                        <Image  Name="ImageAbout" Width="32" Height="32" />
                    </StackPanel>
                </Button>
                <Button VerticalAlignment="Center"   Margin="10,0,0,0" Name="ButtonTwitter">
                    <StackPanel Orientation="Horizontal">
                        <Image  Name="ImageTwitter" Width="32" Height="32" />
                    </StackPanel>
                </Button>


            </ToolBar>

        </Grid>


        <TabControl Name ="TabControl" Margin="5" Grid.Column="0" Grid.Row="1" Grid.RowSpan="2" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <TabItem Name="TabUserOTP" Header="User with OTP">
                    <ListView SelectionMode="Single" Name="ListViewOTPUsers" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <ListView.View>
                <GridView>
                    <GridViewColumn  DisplayMemberBinding="{Binding SamAccountName}" Header="SamAccountName" />
                    <GridViewColumn  DisplayMemberBinding="{Binding GivenName}" Header="Given names" />
                    <GridViewColumn  DisplayMemberBinding="{Binding Surname}" Header="Surename" />
                    <GridViewColumn  DisplayMemberBinding="{Binding UserPrincipalName}" Header="UserPrincipalName" />
                    <GridViewColumn  DisplayMemberBinding="{Binding Mail}" Header="E-Mail" />    
                </GridView>
            </ListView.View>
            </ListView>
            </TabItem>
            <TabItem Name="TabAllUsers" Header="All Users">
                <ListView SelectionMode="Single"  Name="ListViewAllUsers" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn  DisplayMemberBinding="{Binding SamAccountName}" Header="SamAccountName" />
                            <GridViewColumn  DisplayMemberBinding="{Binding GivenName}" Header="Given names" />
                            <GridViewColumn  DisplayMemberBinding="{Binding Surname}" Header="Surename" />
                            <GridViewColumn  DisplayMemberBinding="{Binding UserPrincipalName}" Header="UserPrincipalName" />
                            <GridViewColumn  DisplayMemberBinding="{Binding Mail}" Header="E-Mail" />
                        </GridView>
                    </ListView.View>
                </ListView>
            </TabItem>
        </TabControl>
        <GridSplitter Grid.Column="1" Grid.Row="2" HorizontalAlignment="Left" VerticalAlignment="Stretch" Width="5" Background="Black"/>

        <ListView Margin="5" Name="ListViewDevices" SelectionMode="Single" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="2" Grid.Row="1" Grid.RowSpan="2">
            <ListView.ContextMenu>
                <ContextMenu Name="ContectMenueDevice">
                    <MenuItem Name="AddDevice" Header="Add Device" />
                    <MenuItem Name="ViewDevice" Header="View Device" />
                    <MenuItem Name="RemoveDevice" Header="Remove Device" />
                    <MenuItem Name="RemoveAllDevices" Header="Remove all Devices" />
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn  Width="350"  Header="OTP" />
                </GridView>
            </ListView.View>
        </ListView>

        <!-- StatusBar  Background="LightGreen" Grid.Row="3" Grid.ColumnSpan="3" Margin="5">
            <Label Content="Status: OK"></Label>
        </StatusBar -->

    </Grid>
   
</Window>
'@


function Convert-XAMLtoWindow
{
  param
  (
    [Parameter(Mandatory=$true)]
    [string]
    $XAML
  )
    
  Add-Type -AssemblyName PresentationFramework
    
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  $result = [Windows.Markup.XAMLReader]::Load($reader)
  $reader.Close()
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  while ($reader.Read())
  {
    $name=$reader.GetAttribute('Name')
    if (!$name) { $name=$reader.GetAttribute('x:Name') }
    if($name)
    {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
  }
  $reader.Close()
  $result
}


function Show-WPFWindow
{
  param
  (
    [Parameter(Mandatory)]
    [Windows.Window]
    $Window
  )
    
  $result = $null
  $null = $window.Dispatcher.InvokeAsync{
    $result = $window.ShowDialog()
    Set-Variable -Name result -Value $result -Scope 1
  }.Wait()
  $result
}

function Update-Users{
  param($AllUserTab)

  $win = new-VisualJobWindow -ActionLabel "Update Userlist" -ParentWindow $window
  [string] $filter = $window.TextBoxADFilter.Text
  
  $pool = New-RunspacePool 
     
  $job = New-RunspaceJob -RunspacePool $pool -Arguments $window, $win, $filter,$LDAPServer, $AllUserTab,$AttributeStore,$ODTDefinitionString -Code {
    param( 
      $mainWindow,
      $win1,
      $filter,
      $LDAPServer,
      $AllUserTab,
      $AttributeStore,
      $ODTDefinitionString
    )
    
   
    if($AllUserTab)
    {
      Write-Verbose "Update List"
      $Userlist = @(Get-AdUser -Filter $('SamAccountName -like "' + $filter + '"') -server $LDAPServer -Properties surname,givenname,SamAccountName,Mail,$AttributeStore | Select-Object -Property surname,givenname,SamAccountName,UserPrincipalName,Mail,$AttributeStore)
      
      #$Userlist | Where-Object {$_.mail -ne $null} | % {Add-Member -InputObject $_ -NotePropertyName "EMail" -NotePropertyValue $_.mail -force} 

      $mainWindow.Dispatcher.Invoke([System.Action] { $mainWindow.ListViewAllUsers.ItemsSource = $Userlist  })

    } else
    {
      $ldapFilter = '(&(SamAccountName='+$filter+')('+$AttributeStore+'='+$ODTDefinitionString+'*))'
      #Write-Host  $ldapFilter
      $Userlist  = @((Get-AdUser -LdapFilter $ldapFilter -server $LDAPServer  -Properties surname,givenname,SamAccountName,Mail,$AttributeStore | Select-Object -Property surname,givenname,SamAccountName,UserPrincipalName,Mail,$AttributeStore )) 
      
      #$Userlist | Where-Object {$_.mail -ne $null} | % {Add-Member -InputObject $_ -NotePropertyName "EMail" -NotePropertyValue $_.mail -force} 
          
      $mainWindow.Dispatcher.Invoke([System.Action] { $mainWindow.ListViewOTPUsers.ItemsSource = $Userlist  })
    }

    Write-Verbose "Finished update userlist" -Verbose

    #Close waiting Window

    $win1.Dispatcher.Invoke{$result = $win1.close()}.Wait()

  } -Verbose 

  # Show Window

  Show-VisualJobWindow -window $win
  Clear-Runspace -RunspacePool $pool
}


$window = Convert-XAMLtoWindow -XAML $xaml


#region Define Event Handlers
# Right-Click XAML Text and choose WPF/Attach Events to
# add more handlers

$window.ButtonRefresh.add_Click{
  Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
}

$window.ListViewOTPUsers.add_SelectionChanged{


  if($window.ListViewOTPUsers.ItemsSource -ne $null){

    $window.ListViewOTPUsers.SelectedItem

    [string] $DeviceString = $($window.ListViewOTPUsers.SelectedItem.$AttributeStore) -replace $ODTDefinitionString,''
    $DeviceList = @()

    if(($DeviceString.Length -gt 0) -and ($DeviceString.Substring($DeviceString.Length-1, 1) -eq $ODTSeperator)){
      $DeviceString = $DeviceString.Substring(0, $DeviceString.Length-1)
      $DeviceList = @($DeviceString -split $ODTSeperator)
    }

    $window.ListViewDevices.ItemsSource = $DeviceList
  }
}

$window.ListViewAllUsers.add_SelectionChanged{

  if($window.ListViewAllUsers.ItemsSource -ne $null){

    [string] $DeviceString = $($window.ListViewAllUsers.SelectedItem.$AttributeStore) -replace $ODTDefinitionString,''
    $DeviceList = @()

    if(($DeviceString.Length -gt 0) -and ($DeviceString.Substring($DeviceString.Length-1, 1) -eq $ODTSeperator)){
      $DeviceString = $DeviceString.Substring(0, $DeviceString.Length-1)
      $DeviceList = @($DeviceString -split $ODTSeperator)
    }

    $window.ListViewDevices.ItemsSource = $DeviceList
  }

}

$window.RemoveDevice.add_Click{

  $User = $null
  $DeviceString = $null
  if($window.TabControl.SelectedIndex -eq 0){
    $User = $($window.ListViewOTPUsers.SelectedItem.SamAccountName)
    $DeviceString = $($window.ListViewOTPUsers.SelectedItem.$AttributeStore)

  } else {
    $User = $($window.ListViewAllUsers.SelectedItem.SamAccountName)
    $DeviceString = $($window.ListViewAllUsers.SelectedItem.$AttributeStore)
  }

  $DeviceList = @()

  if(($DeviceString.Length -gt 0) -and ($DeviceString.Substring($DeviceString.Length-1, 1) -eq $ODTSeperator)){
    $DeviceString2 = $DeviceString.Substring(0, $DeviceString.Length-1)
    $DeviceList = @($DeviceString2 -split $ODTSeperator)
  }
  
  if($DeviceList.Count -gt 1){
  
    $DeviceString = $DeviceString -Replace  $( $window.ListViewDevices.SelectedItem + $ODTSeperator) ,""
   
    $Command = 'Set-ADUser -Identity ' + $User + ' -Replace @{'+ $AttributeStore +'="'+$DeviceString+'"}'
    
  } else { #Clear
    $Command = "Set-ADUser -Identity $user -Clear '"+ $AttributeStore +"'"
  }
  
  $result = Out-Message -Message $("SamAccountName = $User `nExecute command : " + $Command) -Type Information -ParentWindow $window -DisableCancel $false
  
  if($result -eq "OK"){
    try{
      Invoke-Expression  $Command -ErrorAction Stop
      $window.ListViewDevices.ItemsSource.Clear()
      $window.ListViewDevices.Items.Refresh()
      Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
        
    } catch {
      Out-Message -Message $($_ | Out-String) -Type Error -ParentWindow $window
    }
  } 
 
}

$window.RemoveAllDevices.add_Click{
  $SamAccountName = $null
  if($window.TabControl.SelectedIndex -eq 0){
    $SamAccountName = $window.ListViewOTPUsers.SelectedItem.SamAccountName
  } else {
    $SamAccountName = $window.ListViewAllUsers.SelectedItem.SamAccountName
  }
  
  $Command = "Set-ADUser -Identity $SamAccountName -Clear '"+ $AttributeStore +"'"
  $result = Out-Message -Message $("SamAccountName = $SamAccountName `nExecute command : " + $Command) -Type Information -ParentWindow $window -DisableCancel $false

  if($result -eq "OK"){
    try{
      Invoke-Expression  $Command -ErrorAction Stop
      $window.ListViewDevices.ItemsSource.Clear()
      $window.ListViewDevices.Items.Refresh()
      
      
      if($window.TabControl.SelectedIndex -eq 0){
        Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
      }
    } catch {
      Out-Message -Message $($_ | Out-String) -Type Information -ParentWindow $window
    }
  } 
}

$window.TabControl.add_SelectionChanged{
  
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.Controls.SelectionChangedEventArgs] $e
  )
  
  #Write-Host $sender
  if($e.Source.name -eq "TabControl"){ #only control changed
    if($window.IsLoaded -and ($window.TabAllUsers.IsSelected)){
      Update-Users -AllUserTab $true
    }
  
    if($window.IsLoaded -and ($window.TabUserOTP.IsSelected)){
      Update-Users -AllUserTab $false
    }
  }
}

$window.ListViewDevices.add_MouseRightButtonDown{

  #Write-Host  $("Mouse" + $window.ListViewDevices.SelectedIndex)
  $window.RemoveDevice.IsEnabled = $false
  $window.AddDevice.IsEnabled = $true
  $window.ViewDevice.IsEnabled = $false
  
  
  if($window.ListViewDevices.Items.Count -eq 0){
    $window.RemoveAllDevices.IsEnabled = $false
    $window.RemoveDevice.IsEnabled = $false
  } else {
    if($window.ListViewDevices.SelectedIndex -ne -1){
    
      $window.RemoveDevice.IsEnabled = $true
      $window.ViewDevice.IsEnabled = $true
    }
    $window.RemoveAllDevices.IsEnabled = $true
  }
    
}


$window.ListViewDevices.add_SelectionChanged{
  if($window.ListViewDevices.Items.Count -eq 0){
    $window.RemoveAllDevices.IsEnabled = $false
    
    $window.RemoveDevice.IsEnabled = $false
    $window.ViewDevice.IsEnabled = $false
    
  } else {
    if($window.ListViewDevices.SelectedIndex -ne -1){
      $window.RemoveDevice.IsEnabled = $true
      
      $window.ViewDevice.IsEnabled = $true
    }
    $window.RemoveAllDevices.IsEnabled = $true
  }
}

#Report Funktion
$window.ButtonExportCSV.add_Click{

  Add-Type -AssemblyName System.Windows.Forms
  $dlg = New-Object System.Windows.Forms.SaveFileDialog
    
  $dlg.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|Excel Worksheet (*.xls)|*.xls|All Files (*.*)|*.*"
  $dlg.SupportMultiDottedExtensions = $true;
  $dlg.InitialDirectory = "$env:Desktop\"
  $dlg.CheckFileExists = $false
    

  if($dlg.ShowDialog() -eq 'Ok'){
    $filter = $window.TextBoxADFilter.Text
    $ldapFilter = '(&(SamAccountName='+$filter+')('+$AttributeStore+'='+$ODTDefinitionString+'*))'
    $Userlist  = @((Get-AdUser -LdapFilter $ldapFilter -server $LDAPServer -Properties surname,givenname,SamAccountName,mail,$AttributeStore )) 
    #Get Devices and Device Count
    foreach($user in $userlist){
      $DeviceList = @()
      [string] $DeviceString = $($user[$AttributeStore]) -replace $ODTDefinitionString,''
      if(($DeviceString.Length -gt 0) -and ($DeviceString.Substring($DeviceString.Length-1, 1) -eq $ODTSeperator)){
        $DeviceString = $DeviceString.Substring(0, $DeviceString.Length-1)
        $DeviceList = @($DeviceString -split $ODTSeperator)
      }
      Add-Member -InputObject $user -MemberType NoteProperty -Name DeviceCount -Value  $DeviceList.Count -Force
    }
    $Userlist | Select-Object -Property surname,givenname,SamAccountName,DeviceCount | Export-Csv -Path $($dlg.FileName) -Encoding UTF8
  }
}

$window.OtpWindow.add_Loaded{

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
}

$window.TextBoxADFilter.add_KeyDown{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
  )
  
  Add-Type -AssemblyName WindowsBase
  if($e.Key -eq [System.Windows.Input.Key]::Return){
  
  
    if($window.TextBoxADFilter.text -eq ""){
      $window.TextBoxADFilter.text = '*'
    }
  
    Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
  }
}

$window.ButtonConfig.add_Click{

  #Change Config
    
  $server = $Global:Configxml.OtpEditConfig.LDAPServer
  $Global:Configxml = New-XMLSettingsDialog -Title "Settings" -SettingsXml $Global:Configxml -EntryKey "OtpEditConfig" -ParentWindow $window
  
  $Global:Configxml.Save($OTPConfigFile)
  Set-GlobalSettings
    
  if($server -ne  $Global:Configxml.OtpEditConfig.LDAPServer){
    Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
  }
}

$window.ButtonTwitter.add_Click{
  New-PSDrive -Name HKCR -PSProvider registry -Root Hkey_Classes_Root | Out-Null
  $browserPath = ((Get-ItemProperty 'HKCR:\http\shell\open\command').'(default)').Split('"')[1]
  & $browserPath "https://twitter.com/nickinformation"
}

$window.ButtonAbout.add_Click{
  New-MessageBox  -Symbol $Global:Shell32Symbols.Star -title "About" -Message $("Andreas Nick - 2020`n") -Hyperlinks `
  @("https://software-Virtualisierung.de", "https://andreasnick.com","https://nick-it.de")  -linksAltText `
  @("My German blog Softwarevirtualisierung`n", "My English blog`n","My Company`n") -EndMessage `
  "`nThanks to Thorsten for the idea and the testing`n Contact: info@nick-it.de" -DisableCancle $true -ParentWindow $window
}


#
#
# Add new Device
#
#

$window.AddDevice.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  
  $User=$null
  if($window.TabControl.SelectedIndex -eq 0){
    $user = $window.ListViewOTPUsers.SelectedItem
    
  } else {
    $user = $window.ListViewAllUsers.SelectedItem
  }

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
    Write-Verbose "Add Device $DeviceString"
  
    $User = $null
    $DeviceString = $null
    if($window.TabControl.SelectedIndex -eq 0){
      $User = $($window.ListViewOTPUsers.SelectedItem.SamAccountName)
      $DeviceString = $($window.ListViewOTPUsers.SelectedItem.$AttributeStore)

    } else {
      $User = $($window.ListViewAllUsers.SelectedItem.SamAccountName)
      $DeviceString = $($window.ListViewAllUsers.SelectedItem.$AttributeStore)
    }
    
    #Device Exist?
    if($DeviceString -notmatch $($Result.Device)){
    
      if($DeviceString -notmatch $('^' + $Global:Configxml.OtpEditConfig.ODTDefinitionString)){
        $DeviceString = $($Global:Configxml.
        OtpEditConfig.ODTDefinitionString + $DeviceString)
      }
    
      $Command = 'Set-ADUser -Identity ' + $User + ' -Replace @{'+ $AttributeStore +'="'+$DeviceString+$NewDevString+'"}'
  
      $res = Out-Message -Message $("SamAccountName = $User `nExecute command : " + $Command) -Type Information -ParentWindow $window -DisableCancel $false
  
      if($res -eq "OK"){
        try{
          Invoke-Expression  $Command -ErrorAction Stop
          $window.ListViewDevices.ItemsSource.Clear()
          $window.ListViewDevices.Items.Refresh()
          Update-Users -AllUserTab ($window.TabControl.SelectedIndex -eq 1)
          
          #
          # Send QRCode
          #

          $res = Out-Message -Message $("Send E-Mail to User $User (" + $Result.EMail +")") -Type Information -ParentWindow $window -DisableCancel $false
          
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


#
#
# Add new Device
#
#

$window.ViewDevice.add_Click{

  $User=$null
  if($window.TabControl.SelectedIndex -eq 0){
    $user = $window.ListViewOTPUsers.SelectedItem
    
  } else {
    $user = $window.ListViewAllUsers.SelectedItem
  }
  
  $EMailAdress = $null
  
  if($User.Mail -eq $null){
    if($User.UserPrincipalName -eq $null){
      $EMailAdress = unknown@unknown.local
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
  
  $User = $null
    
  if($window.TabControl.SelectedIndex -eq 0){
    $User = $($window.ListViewOTPUsers.SelectedItem.SamAccountName)
    
     
  } else {
    $User = $($window.ListViewAllUsers.SelectedItem.SamAccountName)
    
  }

  
  if($window.ListViewDevices.SelectedItem -ne $null){
    $device=$window.ListViewDevices.SelectedItem 
    #$text='phonexx=T7UD3QIHEVDLVPGSCNUYAXX2I4&' 
    $pattern = '^(.*)=([A-Z|0-9].*)&$'
    
    $res=[RegEx]::Matches($device, $pattern)
    
    $res.Value
    if($res.success){
      $Secret = $res.Groups[2].Value
      $DeviceName = $res.Groups[1].Value
      #Write-Verbose $Secret
      
      $result = New-OtpQRCodeWindow -Secret $Secret -Settings $Global:Configxml -Parent $window -Email $EMailAdress -UserPrincipalName $username -Device $DeviceName -CanEdit $False

      if($result -ne $Null){
        try
        {
          $res = Out-Message -Message $("Send E-Mail to User $User (" + $Result.EMail +")") -Type Information -ParentWindow $window -DisableCancel $false
          if($res -eq "OK"){
          
            [int] $Port = $($Global:Configxml.OtpEditConfig.SMTPPort.InnerText)

            Send-QRCodeEMail -SMTPPort $Port  -SMPTServer $($Global:Configxml.OtpEditConfig.SMTPServer.InnerText) `
            -From $($Global:Configxml.OtpEditConfig.SMTPMailFrom.InnerText) `
            -Subject $($Global:Configxml.OtpEditConfig.SMTPSubject.InnerText) -SMTPUseSSL ($Global:Configxml.OtpEditConfig.SMTPUseSSL.InnerText -eq "true") `
            -SMTPMailuser $($Global:Configxml.OtpEditConfig.SMTPUser.InnerText) -SMTPMailPassword $($Global:Configxml.OtpEditConfig.SMTPPassword.InnerText) `
            -To $Result.EMail -QRCode $result.QRCode -UserName $User
          }
        }

        catch {
          Out-Message -Message $($_ | Out-String) -Type Error -ParentWindow $window
        }
      }
    }
  
    else {
      Write-Verbose "No Device selected"
    }
  }
}


$window.add_Loaded{
  Update-Users -AllUserTab $false
}

#endregion Event Handlers
#region Manipulate Window Content
$window.ImageRefresh.Source = Get-ImageSourceFromShell32dll -IconIndex $Shell32Symbols.Refresh
$window.ImageConfig.Source = Get-ImageSourceFromShell32dll -IconIndex $Shell32Symbols.Gear
$window.ImageAbout.Source = Get-ImageSourceFromShell32dll -IconIndex $Shell32Symbols.Star
$Window.ImageTwitter.Source =  Convert-Base64ToImageSourtce -bgImage64 (Get-TwitterBitmapBase64)
$window.ButtonTwitter.ToolTip = "For contact and infos: @nickinformation"

$window.icon = $window.ImageConfig.Source = Get-ImageSourceFromShell32dll -IconIndex $Shell32Symbols.Gear

#endregion

$Result = Show-WPFWindow -Window $window
