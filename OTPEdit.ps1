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
   
    $Command = 'Set-ADUser -Identity ' + $User + ' -Replace @{'''+ $AttributeStore +'''="'+$DeviceString+'"}'
    
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
    
      $Command = 'Set-ADUser -Identity ' + $User + ' -Replace @{'''+ $AttributeStore +'''="'+$DeviceString+$NewDevString+'"}'
  
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
            #-To $Result.EMail -QRCode $result.QRCode -UserName $User -Secret $Result.Secret -DeviceName $Result.Device -UserPrincipleName $Result.UserPrincipleName
            start-steroids

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

# SIG # Begin signature block
# MIIm6AYJKoZIhvcNAQcCoIIm2TCCJtUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJzWhw3rJW17zYKuE/ycnHkP8
# xhmggh/QMIIFgTCCBGmgAwIBAgIQOXJEOvkit1HX02wQ3TE1lTANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTE5MDMxMjAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcg
# SmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJU
# UlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENlcnRpZmljYXRp
# b24gQXV0aG9yaXR5MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAgBJl
# FzYOw9sIs9CsVw127c0n00ytUINh4qogTQktZAnczomfzD2p7PbPwdzx07HWezco
# EStH2jnGvDoZtF+mvX2do2NCtnbyqTsrkfjib9DsFiCQCT7i6HTJGLSR1GJk23+j
# BvGIGGqQIjy8/hPwhxR79uQfjtTkUcYRZ0YIUcuGFFQ/vDP+fmyc/xadGL1RjjWm
# p2bIcmfbIWax1Jt4A8BQOujM8Ny8nkz+rwWWNR9XWrf/zvk9tyy29lTdyOcSOk2u
# TIq3XJq0tyA9yn8iNK5+O2hmAUTnAU5GU5szYPeUvlM3kHND8zLDU+/bqv50TmnH
# a4xgk97Exwzf4TKuzJM7UXiVZ4vuPVb+DNBpDxsP8yUmazNt925H+nND5X4OpWax
# KXwyhGNVicQNwZNUMBkTrNN9N6frXTpsNVzbQdcS2qlJC9/YgIoJk2KOtWbPJYjN
# hLixP6Q5D9kCnusSTJV882sFqV4Wg8y4Z+LoE53MW4LTTLPtW//e5XOsIzstAL81
# VXQJSdhJWBp/kjbmUZIO8yZ9HE0XvMnsQybQv0FfQKlERPSZ51eHnlAfV1SoPv10
# Yy+xUGUJ5lhCLkMaTLTwJUdZ+gQek9QmRkpQgbLevni3/GcV4clXhB4PY9bpYrrW
# X1Uu6lzGKAgEJTm4Diup8kyXHAc/DVL17e8vgg8CAwEAAaOB8jCB7zAfBgNVHSME
# GDAWgBSgEQojPpbxB+zirynvgqV/0DCktDAdBgNVHQ4EFgQUU3m/WqorSs9UgOHY
# m8Cd8rIDZsswDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEQYDVR0g
# BAowCDAGBgRVHSAAMEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2Rv
# Y2EuY29tL0FBQUNlcnRpZmljYXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgw
# JjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3
# DQEBDAUAA4IBAQAYh1HcdCE9nIrgJ7cz0C7M7PDmy14R3iJvm3WOnnL+5Nb+qh+c
# li3vA0p+rvSNb3I8QzvAP+u431yqqcau8vzY7qN7Q/aGNnwU4M309z/+3ri0ivCR
# lv79Q2R+/czSAaF9ffgZGclCKxO/WIu6pKJmBHaIkU4MiRTOok3JMrO66BQavHHx
# W/BBC5gACiIDEOUMsfnNkjcZ7Tvx5Dq2+UUTJnWvu6rvP3t3O9LEApE9GQDTF1w5
# 2z97GA1FzZOFli9d31kWTz9RvdVFGD/tSo7oBmF0Ixa1DVBzJ0RHfxBdiSprhTEU
# xOipakyAvGp4z7h/jnZymQyd/teRCBaho1+VMIIF9TCCA92gAwIBAgIQHaJIMG+b
# JhjQguCWfTPTajANBgkqhkiG9w0BAQwFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCk5ldyBKZXJzZXkxFDASBgNVBAcTC0plcnNleSBDaXR5MR4wHAYDVQQKExVU
# aGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBSU0EgQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMTgxMTAyMDAwMDAwWhcNMzAxMjMxMjM1
# OTU5WjB8MQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxJDAi
# BgNVBAMTG1NlY3RpZ28gUlNBIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAIYijTKFehifSfCWL2MIHi3cfJ8Uz+MmtiVmKUCG
# VEZ0MWLFEO2yhyemmcuVMMBW9aR1xqkOUGKlUZEQauBLYq798PgYrKf/7i4zIPoM
# GYmobHutAMNhodxpZW0fbieW15dRhqb0J+V8aouVHltg1X7XFpKcAC9o95ftanK+
# ODtj3o+/bkxBXRIgCFnoOc2P0tbPBrRXBbZOoT5Xax+YvMRi1hsLjcdmG0qfnYHE
# ckC14l/vC0X/o84Xpi1VsLewvFRqnbyNVlPG8Lp5UEks9wO5/i9lNfIi6iwHr0bZ
# +UYc3Ix8cSjz/qfGFN1VkW6KEQ3fBiSVfQ+noXw62oY1YdMCAwEAAaOCAWQwggFg
# MB8GA1UdIwQYMBaAFFN5v1qqK0rPVIDh2JvAnfKyA2bLMB0GA1UdDgQWBBQO4Tqo
# Uzox1Yq+wbutZxoDha00DjAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHSUEFjAUBggrBgEFBQcDAwYIKwYBBQUHAwgwEQYDVR0gBAowCDAG
# BgRVHSAAMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNv
# bS9VU0VSVHJ1c3RSU0FDZXJ0aWZpY2F0aW9uQXV0aG9yaXR5LmNybDB2BggrBgEF
# BQcBAQRqMGgwPwYIKwYBBQUHMAKGM2h0dHA6Ly9jcnQudXNlcnRydXN0LmNvbS9V
# U0VSVHJ1c3RSU0FBZGRUcnVzdENBLmNydDAlBggrBgEFBQcwAYYZaHR0cDovL29j
# c3AudXNlcnRydXN0LmNvbTANBgkqhkiG9w0BAQwFAAOCAgEATWNQ7Uc0SmGk295q
# Koyb8QAAHh1iezrXMsL2s+Bjs/thAIiaG20QBwRPvrjqiXgi6w9G7PNGXkBGiRL0
# C3danCpBOvzW9Ovn9xWVM8Ohgyi33i/klPeFM4MtSkBIv5rCT0qxjyT0s4E307dk
# sKYjalloUkJf/wTr4XRleQj1qZPea3FAmZa6ePG5yOLDCBaxq2NayBWAbXReSnV+
# pbjDbLXP30p5h1zHQE1jNfYw08+1Cg4LBH+gS667o6XQhACTPlNdNKUANWlsvp8g
# JRANGftQkGG+OY96jk32nw4e/gdREmaDJhlIlc5KycF/8zoFm/lv34h/wCOe0h5D
# ekUxwZxNqfBZslkZ6GqNKQQCd3xLS81wvjqyVVp4Pry7bwMQJXcVNIr5NsxDkuS6
# T/FikyglVyn7URnHoSVAaoRXxrKdsbwcCtp8Z359LukoTBh+xHsxQXGaSynsCz1X
# UNLK3f2eBVHlRHjdAd6xdZgNVCT98E7j4viDvXK6yz067vBeF5Jobchh+abxKgoL
# pbn0nu6YMgWFnuv5gynTxix9vTp3Los3QqBqgu07SqqUEKThDfgXxbZaeTMYkuO1
# dfih6Y4KJR7kHvGfWocj/5+kUZ77OYARzdu1xKeogG/lU9Tg46LC0lsa+jImLWpX
# cBw8pFguo/NbSwfcMlnzh6cabVgwggZkMIIFTKADAgECAhAOWThJHf2EAGPmiQKS
# aTawMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVh
# dGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEkMCIGA1UEAxMbU2VjdGlnbyBSU0EgQ29kZSBTaWduaW5nIENB
# MB4XDTIwMDgxNzAwMDAwMFoXDTIzMDgxNzIzNTk1OVowga0xCzAJBgNVBAYTAkRF
# MQ4wDAYDVQQRDAUzMDUzOTEWMBQGA1UECAwNTmllZGVyc2FjaHNlbjERMA8GA1UE
# BwwISGFubm92ZXIxEzARBgNVBAkMCkRyaWJ1c2NoIDIxJjAkBgNVBAoMHU5pY2sg
# SW5mb3JtYXRpb25zdGVjaG5payBHbWJIMSYwJAYDVQQDDB1OaWNrIEluZm9ybWF0
# aW9uc3RlY2huaWsgR21iSDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# ANperdsEZyxEKnjFFen+7kV+EfL5NpZUz6Yl4YhkpYYPRHf7z9yBWGvc0cjVlcB4
# Zr9AhqcVPns5EqUjT7TCOJxuGjScN+6vTt1KOxrgOjMlvoUztKrKbOsGsdhL5OhU
# ANOZ2vwOvc0lQM1cMCQsW//iVnsu1noYiC1ju42tTD9yciIiSIC3kfL1mJKBFzW3
# Y0t7tdNyIES5RtmE0KeqaHJBtbA3sbubY1BB/TxTWVTXNjr2HuvsbNuyTUd84C3H
# Hgoed7hrSWv07fZvvDF51Nnn/wZrRU2wtHE/HJfZ+btgctI5PQsxmInBoxPTgL5i
# MuKzfU6Vk04vymc3a6ABXuXfUSUB5OcPZCnan5V8Qa0nK2l+KCD3aW+ZqvZax+F0
# I/ij1MFxtqFgLqKaebOlri8R9Wv3hOhkfGoDD+DNizhQDeznJDQES3c1Bu6FgKl7
# LRdVfaM/qwSq6s817gbOcPGuOe4zkue4vBzrzvKfPceptxtCIgNF/3fQxSick1Gn
# tuhPCzW8i7OFUoDK4PY6jdZtP3tjb3oJQym0Fjs0p1g1BMLU4d31FD+KGeNOO37n
# 8KzVPPm0FZf5zSd/3NeRRZy0fI4ZCJvJMJQguSrMeTjxfLFnTR5m6Na0SBUBLsXd
# ozdWvwhZdxJNt4tYkPLUZUQLR7oSFXOR25ZWm7Wdf8GlAgMBAAGjggGuMIIBqjAf
# BgNVHSMEGDAWgBQO4TqoUzox1Yq+wbutZxoDha00DjAdBgNVHQ4EFgQUq0Q7xQIR
# CzrWxe5EAkrwXkJJi9owDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEoGA1UdIARDMEEw
# NQYMKwYBBAGyMQECAQMCMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8vc2VjdGlnby5j
# b20vQ1BTMAgGBmeBDAEEATBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsLnNl
# Y3RpZ28uY29tL1NlY3RpZ29SU0FDb2RlU2lnbmluZ0NBLmNybDBzBggrBgEFBQcB
# AQRnMGUwPgYIKwYBBQUHMAKGMmh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGln
# b1JTQUNvZGVTaWduaW5nQ0EuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5z
# ZWN0aWdvLmNvbTAcBgNVHREEFTATgRFhLm5pY2tAbmljay1pdC5kZTANBgkqhkiG
# 9w0BAQsFAAOCAQEAC6KzO/xsS5EkS5KLY873pwHamFGMpOzIEHeoAiqpX7LMy8Gu
# 41Rznsou/ZQGjYS9HpgezYk4kk5AoNGY/+ObcnSIvNFOr7EkYwZt+uTejdzgbY8g
# hDmhdm3XpfjqO+DjJe6xtf8Qfies7bnXhcKvNFTvycIPjikVvxF/tbfFzx9iRqzO
# XCgznnMR2e+VbK6vjZJeMgR0KaHRmXOfXt/jX7cJt1j0TateKpECeFxawKA8fmle
# ZBw2KVXEI+0lDmuxk4oRRkLdH0nUv06mqicgsfwUFbCkU26tgGJemGdmtf5VcVM9
# 03K/4xJkkhuOlwMA1e9XU/is0kSk/KW1G9msvDCCBuwwggTUoAMCAQICEDAPb6zd
# Zph0fKlGNqd4LbkwDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UEChMV
# VGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENl
# cnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTE5MDUwMjAwMDAwMFoXDTM4MDExODIz
# NTk1OVowfTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3Rl
# cjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSUw
# IwYDVQQDExxTZWN0aWdvIFJTQSBUaW1lIFN0YW1waW5nIENBMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAyBsBr9ksfoiZfQGYPyCQvZyAIVSTuc+gPlPv
# s1rAdtYaBKXOR4O168TMSTTL80VlufmnZBYmCfvVMlJ5LsljwhObtoY/AQWSZm8h
# q9VxEHmH9EYqzcRaydvXXUlNclYP3MnjU5g6Kh78zlhJ07/zObu5pCNCrNAVw3+e
# olzXOPEWsnDTo8Tfs8VyrC4Kd/wNlFK3/B+VcyQ9ASi8Dw1Ps5EBjm6dJ3VV0Rc7
# NCF7lwGUr3+Az9ERCleEyX9W4L1GnIK+lJ2/tCCwYH64TfUNP9vQ6oWMilZx0S2U
# TMiMPNMUopy9Jv/TUyDHYGmbWApU9AXn/TGs+ciFF8e4KRmkKS9G493bkV+fPzY+
# DjBnK0a3Na+WvtpMYMyou58NFNQYxDCYdIIhz2JWtSFzEh79qsoIWId3pBXrGVX/
# 0DlULSbuRRo6b83XhPDX8CjFT2SDAtT74t7xvAIo9G3aJ4oG0paH3uhrDvBbfel2
# aZMgHEqXLHcZK5OVmJyXnuuOwXhWxkQl3wYSmgYtnwNe/YOiU2fKsfqNoWTJiJJZ
# y6hGwMnypv99V9sSdvqKQSTUG/xypRSi1K1DHKRJi0E5FAMeKfobpSKupcNNgtCN
# 2mu32/cYQFdz8HGj+0p9RTbB942C+rnJDVOAffq2OVgy728YUInXT50zvRq1naHe
# lUF6p4MCAwEAAaOCAVowggFWMB8GA1UdIwQYMBaAFFN5v1qqK0rPVIDh2JvAnfKy
# A2bLMB0GA1UdDgQWBBQaofhhGSAPw0F3RSiO0TVfBhIEVTAOBgNVHQ8BAf8EBAMC
# AYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcDCDARBgNV
# HSAECjAIMAYGBFUdIAAwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC51c2Vy
# dHJ1c3QuY29tL1VTRVJUcnVzdFJTQUNlcnRpZmljYXRpb25BdXRob3JpdHkuY3Js
# MHYGCCsGAQUFBwEBBGowaDA/BggrBgEFBQcwAoYzaHR0cDovL2NydC51c2VydHJ1
# c3QuY29tL1VTRVJUcnVzdFJTQUFkZFRydXN0Q0EuY3J0MCUGCCsGAQUFBzABhhlo
# dHRwOi8vb2NzcC51c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4ICAQBtVIGl
# M10W4bVTgZF13wN6MgstJYQRsrDbKn0qBfW8Oyf0WqC5SVmQKWxhy7VQ2+J9+Z8A
# 70DDrdPi5Fb5WEHP8ULlEH3/sHQfj8ZcCfkzXuqgHCZYXPO0EQ/V1cPivNVYeL9I
# duFEZ22PsEMQD43k+ThivxMBxYWjTMXMslMwlaTW9JZWCLjNXH8Blr5yUmo7Qjd8
# Fng5k5OUm7Hcsm1BbWfNyW+QPX9FcsEbI9bCVYRm5LPFZgb289ZLXq2jK0KKIZL+
# qG9aJXBigXNjXqC72NzXStM9r4MGOBIdJIct5PwC1j53BLwENrXnd8ucLo0jGLmj
# wkcd8F3WoXNXBWiap8k3ZR2+6rzYQoNDBaWLpgn/0aGUpk6qPQn1BWy30mRa2Coi
# wkud8TleTN5IPZs0lpoJX47997FSkc4/ifYcobWpdR9xv1tDXWU9UIFuq/DQ0/yy
# sx+2mZYm9Dx5i1xkzM3uJ5rloMAMcofBbk1a0x7q8ETmMm8c6xdOlMN4ZSA7D0Gq
# H+mhQZ3+sbigZSo04N6o+TzmwTC7wKBjLPxcFgCo0MR/6hGdHgbGpm0yXbQ4CStJ
# B6r97DDa8acvz7f9+tCjhNknnvsBZne5VhDhIG7GrrH5trrINV0zdo7xfCAMKneu
# taIChrop7rRaALGMq+P5CslUXdS5anSevUiumDCCBvYwggTeoAMCAQICEQCQOX+a
# 0ko6E/K9kV8IOKlDMA0GCSqGSIb3DQEBDAUAMH0xCzAJBgNVBAYTAkdCMRswGQYD
# VQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UEAxMcU2VjdGlnbyBSU0EgVGltZSBT
# dGFtcGluZyBDQTAeFw0yMjA1MTEwMDAwMDBaFw0zMzA4MTAyMzU5NTlaMGoxCzAJ
# BgNVBAYTAkdCMRMwEQYDVQQIEwpNYW5jaGVzdGVyMRgwFgYDVQQKEw9TZWN0aWdv
# IExpbWl0ZWQxLDAqBgNVBAMMI1NlY3RpZ28gUlNBIFRpbWUgU3RhbXBpbmcgU2ln
# bmVyICMzMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAkLJxP3nh1LmK
# F8zDl8KQlHLtWjpvAUN/c1oonyR8oDVABvqUrwqhg7YT5EsVBl5qiiA0cXu7Ja0/
# WwqkHy9sfS5hUdCMWTc+pl3xHl2AttgfYOPNEmqIH8b+GMuTQ1Z6x84D1gBkKFYi
# sUsZ0vCWyUQfOV2csJbtWkmNfnLkQ2t/yaA/bEqt1QBPvQq4g8W9mCwHdgFwRd7D
# 8EJp6v8mzANEHxYo4Wp0tpxF+rY6zpTRH72MZar9/MM86A2cOGbV/H0em1mMkVpC
# V1VQFg1LdHLuoCox/CYCNPlkG1n94zrU6LhBKXQBPw3gE3crETz7Pc3Q5+GXW1X3
# KgNt1c1i2s6cHvzqcH3mfUtozlopYdOgXCWzpSdoo1j99S1ryl9kx2soDNqseEHe
# ku8Pxeyr3y1vGlRRbDOzjVlg59/oFyKjeUFiz/x785LaruA8Tw9azG7fH7wir7c4
# EJo0pwv//h1epPPuFjgrP6x2lEGdZB36gP0A4f74OtTDXrtpTXKZ5fEyLVH6Ya1N
# 6iaObfypSJg+8kYNabG3bvQF20EFxhjAUOT4rf6sY2FHkbxGtUZTbMX04YYnk4Q5
# bHXgHQx6WYsuy/RkLEJH9FRYhTflx2mn0iWLlr/GreC9sTf3H99Ce6rrHOnrPVrd
# +NKQ1UmaOh2DGld/HAHCzhx9zPuWFcUCAwEAAaOCAYIwggF+MB8GA1UdIwQYMBaA
# FBqh+GEZIA/DQXdFKI7RNV8GEgRVMB0GA1UdDgQWBBQlLmg8a5orJBSpH6LfJjrP
# FKbx4DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMDUGDCsGAQQBsjEBAgEDCDAlMCMGCCsGAQUF
# BwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQUzAIBgZngQwBBAIwRAYDVR0fBD0w
# OzA5oDegNYYzaHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUlNBVGltZVN0
# YW1waW5nQ0EuY3JsMHQGCCsGAQUFBwEBBGgwZjA/BggrBgEFBQcwAoYzaHR0cDov
# L2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUlNBVGltZVN0YW1waW5nQ0EuY3J0MCMG
# CCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwF
# AAOCAgEAc9rtaHLLwrlAoTG7tAOjLRR7JOe0WxV9qOn9rdGSDXw9NqBp2fOaMNqs
# adZ0VyQ/fg882fXDeSVsJuiNaJPO8XeJOX+oBAXaNMMU6p8IVKv/xH6WbCvTlOu0
# bOBFTSyy9zs7WrXB+9eJdW2YcnL29wco89Oy0OsZvhUseO/NRaAA5PgEdrtXxZC+
# d1SQdJ4LT03EqhOPl68BNSvLmxF46fL5iQQ8TuOCEmLrtEQMdUHCDzS4iJ3IIvET
# atsYL254rcQFtOiECJMH+X2D/miYNOR35bHOjJRs2wNtKAVHfpsu8GT726QDMRB8
# Gvs8GYDRC3C5VV9HvjlkzrfaI1Qy40ayMtjSKYbJFV2Ala8C+7TRLp04fDXgDxzt
# G0dInCJqVYLZ8roIZQPl8SnzSIoJAUymefKithqZlOuXKOG+fRuhfO1WgKb0IjOQ
# 5IRT/Cr6wKeXqOq1jXrO5OBLoTOrC3ag1WkWt45mv1/6H8Sof6ehSBSRDYL8vU2Z
# 7cnmbDb+d0OZuGktfGEv7aOwSf5bvmkkkf+T/FdpkkvZBT9thnLTotDAZNI6QsEa
# A/vQ7ZohuD+vprJRVNVMxcofEo1XxjntXP/snyZ2rWRmZ+iqMODSrbd9sWpBJ24D
# iqN04IoJgm6/4/a3vJ4LKRhogaGcP24WWUsUCQma5q6/YBXdhvUxggaCMIIGfgIB
# ATCBkDB8MQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxJDAi
# BgNVBAMTG1NlY3RpZ28gUlNBIENvZGUgU2lnbmluZyBDQQIQDlk4SR39hABj5okC
# kmk2sDAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUjhBT9vAISTQAs3cTieNODPQnLXswDQYJKoZI
# hvcNAQEBBQAEggIAcg5bL+GiyaS8OKDaeuPR34C4uKwwrlWDBttqz8lzfBqzHsEB
# X5xoJHDRvhuJeO+H5O8araHFGParD8zbaQ1YAwJ0u16adHRuxMOXFoIoeJYpEXjv
# 5AgGY/kdxlNwuozFIiP9671LM/ZC5YNfporTXBKUB+n1+k9pwoGy+tNIU2bDu2a1
# 5RuNJXzBH3zrptOvoAXDpAED8urcS1az4q8WPpZteVz93sUdyAvncE1JhKEHNdYJ
# j/o+CxMluz1TNUS5fDhKTajrDZ8H/23Nh67Mt8BBV0LQc4wNFexpwtevscLX4wzj
# 2DRDW5gdhiBvIkOu7+WO1Ya3jnVAFdVTYwQz56XuGMiFbiKVPZcAeaIhxXlYGq29
# Yj/2nzPr7mrIBawMxzqcepzMll6J5Bc14SPolNU3jxBQI/2Mu6fWnLKAyg4/RcZi
# F+rn3b6uYyw94rStnE2uwZ7oId76SuDVxiQbEX0IKFmT2W0XY1zcSCshCtb/Bdpe
# yRo6XQXRHGRLlGMX08kIYjSDIcRPuEmsh4Tp2qEsvD47ocaSYOWFJhegAPrQ4LGJ
# RsitEtT+yfZnxIfXz7tp910fsmXbOs1EL1BBJ4/pTuuEkBM6ML5OO5Ssfsi8zZ/n
# iwuqEhM2kgfk3LoEffqAX3C02dfSl/Y90EkmLtPPg5+j5W4KmXVfRKBWMlGhggNM
# MIIDSAYJKoZIhvcNAQkGMYIDOTCCAzUCAQEwgZIwfTELMAkGA1UEBhMCR0IxGzAZ
# BgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYG
# A1UEChMPU2VjdGlnbyBMaW1pdGVkMSUwIwYDVQQDExxTZWN0aWdvIFJTQSBUaW1l
# IFN0YW1waW5nIENBAhEAkDl/mtJKOhPyvZFfCDipQzANBglghkgBZQMEAgIFAKB5
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIyMDYx
# MDE2NDgwNVowPwYJKoZIhvcNAQkEMTIEMA3vKHfr9++TON7wOPt7Nc9Zrm+uEy6G
# 6sAWDOA3RPSVFzqhciaCnRUEwMyhyhiYDTANBgkqhkiG9w0BAQEFAASCAgArkSSY
# ji1ARxIwpNIM0B635XqSozGwji05wcqZp2Gw1nItCNvHpxnoQIX6miH8byLt/Rkb
# GoG43xXOTAPsvCQKLp8rnwCCUd+c3euKu82iLkbRUJJc2iWmiQle/cOqq+Eac0on
# ynZXhlOczDUbBicQ4PCtNQEmD83hnWmexZKyVGoEA1Uy/zgIR/Z7vV+xU74/OQ1L
# H7FF4Kf5IDNQ8ePLm8LCQF2XvbcmKOJONyLVr0zWVd9fGxduak4b5FGBU3UhpnPH
# SU5SgqHsovwNzp7LUFIULgyotZdoHZdqkgXY5tmQwoNGjSpd4QKaSVgEg8MLdLMp
# ALETQiWjWAfNZzzLOZKTsDNCfGz/dV1BkDssiy8Ej8eIUpqO9oBvF8t5dKqO92nQ
# dfLtnfwosnObT+dF5MW60kRQ/oaaIWCcwNHbdCnqJTpzTNvWuUPQyNV/3N408EiB
# x1ZSBALHi6SB1q7S1L1xiCwxwu4uT32MYferyzkF1rBM+rSaIlSQG/01/6cxH/Ao
# vRntnCy1CEzXXB4ZEwmORDkK65/jRcy3ymg2ArgvVzRCXDPRfOHTtsOLO+TgqZim
# cQiJ0N9crGoXLg+rH5k3Wq/X9bXRSErCsCDqRYAcOanK10jrcXwDbFxnmYgZO/N3
# reSoXl1EtIHzifs7m/4CeGcbjjo/OuIg+L2N8A==
# SIG # End signature block
