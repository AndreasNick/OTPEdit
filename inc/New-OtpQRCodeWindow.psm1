
function New-OtpQRCodeWindow
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][String] $Secret,
    [Parameter(Mandatory=$true)][xml] $Settings,
    [String] $Device = "Provisioned",
    [String] $Email = "Urantest@uran.local",
    [String] $UserPrincipalName = "Urantest@uran.local",
    $ParentWindow = $null, #Center Parent
    [bool] $CanEdit = $true
  )
  
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName PresentationCore
  $xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   MinWidth="470"
   MinHeight="320" 
   Width ="470"
   SizeToContent="Height"
   Title="Ctx-OTP-Code-Email"
   Topmost="True">
   <Grid Margin="10,10,10,10">
      <Grid.ColumnDefinitions>
         <ColumnDefinition Width="Auto"/>
         <ColumnDefinition Width="100"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="5"/>
        </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

        <StackPanel Orientation="Vertical" Grid.Column="0" Grid.Row="0">
            <TextBlock  Margin="5">QR-Code:</TextBlock>
            <Image Name="ImageQECode" Height="180" Width="180"  Source="C:/Users/Andreas/Desktop/qr.png" Grid.RowSpan="3" Margin="5" />
        </StackPanel>
       
        <StackPanel Margin="5" Orientation="Vertical" Grid.Column="1" Grid.Row="0">
            <TextBlock Margin="5">User Name</TextBlock>
            <TextBlock Margin="5">Email</TextBlock>
            <TextBlock Margin="5">DeviceName</TextBlock>
            <TextBlock Margin="5">Secret</TextBlock>
        </StackPanel>

        <StackPanel Margin="5" Orientation="Vertical" Grid.Column="2" Grid.Row="0">
            <TextBox Margin="5" Name="TxtName">test</TextBox>
            <TextBox Margin="5" Name="TxtEmail">test</TextBox>
            <TextBox Margin="5" Name="TxtDeviceName">Provisioned</TextBox>
            <Border Margin="5"  Background="GhostWhite" BorderBrush="Gainsboro" BorderThickness="1">
                <TextBox  IsReadOnly="True" TextWrapping="Wrap"  Name="TxtSecret">abc</TextBox>
            </Border>
        </StackPanel>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,10,0,0" Grid.Row="4" Grid.ColumnSpan="3">
        <Button Name="ButSave" MinWidth="80" Height="22" Margin="5">SaveQR</Button>
        <Button Name="ButSend" MinWidth="80" Height="22" Margin="5">Send</Button>
        <Button Name="ButCancel" MinWidth="80" Height="22" Margin="5">Cancel</Button>
      </StackPanel>
   </Grid>
</Window>


'@
  #endregion

  #region Code Behind
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
      [Parameter(Mandatory=$true)]
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
  
  function Set-WindowsData{

    $ba = $Bitmap = Get-GenerateCtxOTPQR -secret   $window.TxtSecret.Text -Width 5 -verbose -DeviceName $Window.TxtDeviceName.text -UserPrincipleName $window.TxtName.Text
    $st =  new-object System.Io.MemoryStream(,$ba)
    $st.Position = 0
    $bitmapimage = new-Object System.Windows.Media.Imaging.BitmapImage
    $bitmapimage.BeginInit()
    $bitmapimage.StreamSource = $st
    $bitmapimage.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmapimage.EndInit()
    $window.ImageQECode.source = $bitmapimage 
  }
  
  #endregion Code Behind

  #region Convert XAML to Window
  $window = Convert-XAMLtoWindow -XAML $xaml 
  #endregion

  #region Define Event Handlers
  # Right-Click XAML Text and choose WPF/Attach Events to
  # add more handlers
  $window.ButCancel.add_Click(
    {
      $window.DialogResult = $false
      
    }
  )

  $window.ButSend.add_Click(
    {
    
      $window.DialogResult = $true
    
    }
    
  )
  $window.TxtName.add_LostFocus{
    Set-WindowsData
  }
  $window.TxtDeviceName.add_LostFocus{
    Set-WindowsData
  }
  $window.ButSave.add_Click{


    $ba =  Get-GenerateCtxOTPQR -secret $window.TxtSecret.Text -Width 5 -verbose -DeviceName $Window.TxtDeviceName.text -UserPrincipleName $window.TxtName.Text 
  
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    
    $dlg.Filter = "PNG Files (*.png)|*.png"# |jpg Files (*.jpg)|*.jpg|bmp File (*.bmp)|*.bmp"
    $dlg.SupportMultiDottedExtensions = $true;
    $dlg.InitialDirectory = "$env:Desktop\"
    $dlg.CheckFileExists = $false

    if($dlg.ShowDialog() -eq 'Ok'){
      [System.IO.File]::WriteAllBytes($dlg.FileName, $ba)
    }
  }
  
  #endregion Event Handlers

  #region Manipulate Window Content
  #Center to Parent
  if($ParentWindow -ne $null){
    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $Window.Owner = $ParentWindow
  } else {
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
  }
  
  $window.TxtName.Text = $UserPrincipalName
  $window.Icon = Get-ImageSourceFromShell32dll -IconIndex $Global:Shell32Symbols.UsersAndKey 
  
  $Script:sec = $null
    
  if($Secret -eq ""){
    $pass = GetHexPass
    $sec = HexToBase32 -HexString $pass
  } else {
    $sec = $Secret
  }
  
  $window.TxtSecret.Text = $Sec
  $Window.TxtDeviceName.text = $Device
  $window.TxtEmail.text = $Email
    
  $null = $window.TxtName.Focus()
     
  Set-WindowsData
  
  $window.TxtDeviceName.IsEnabled = $CanEdit
  $window.TxtName.IsEnabled = $CanEdit
  
  #endregion

  # Show Window
  $result = Show-WPFWindow -Window $window

  #endregion Process results

  if ($result -eq $true)
  {
    [PSCustomObject]@{
      QRCode = Get-GenerateCtxOTPQR -secret $window.TxtSecret.Text -Width 5 -verbose -DeviceName $Window.TxtDeviceName.text -UserPrincipleName $window.TxtName.Text 
      Device = $Window.TxtDeviceName.text 
      Secret = $window.TxtSecret.Text
      UserPrincipleName = $window.TxtName.Text 
      EMail = $window.TxtEmail.text
    }
  }

  else
  {
    Return $Null
  
  }

}

