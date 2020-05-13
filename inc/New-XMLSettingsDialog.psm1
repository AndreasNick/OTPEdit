
function New-XMLSettingsDialog
{
  <#
      .SYNOPSIS
      Describe purpose of "Get-SettingsDialog" in 1-2 sentences.

      .DESCRIPTION
      Create a WPF Dialog for a Settings XML file 

      .PARAMETER Title
      Title of the window

      .PARAMETER SettingsXml
      a simle xml element

      $SettingsXml = [xml] @'
      <Settings>
      <SMTPServer Type="String">smtp.gmail.com</SMTPServer>
      <SMTPUser Type="String"></SMTPUser>
      <SMTPPassword Type="Password"></SMTPPassword>
      <SMTPPort Type="String"></SMTPPort>
      <SMTPMailFrom Type="String">Homer@Nick-It.de</SMTPMailFrom>
      <SMTPUseSSL Type="Switch">true</SMTPUseSSL>
      <SMTPSubject Type="String"></SMTPSubject>
      </Settings>
      '@

      .PARAMETER EntryKey
      Entry Key is the root string of the xml

      .PARAMETER ParentWindow
      Parent WPF Window to center

      .EXAMPLE
    
      New-SettingsDialog -Settings $SettingsXml -EntryKey "Settings" -Title "OTP-Config" 
    
      .NOTES
    

      .LINK
    

      .INPUTS
    

      .OUTPUTS
      a new XML
  #>



  param
  (
    [String] $Title = "Settings",
    [xml] $SettingsXml,
    [String] $EntryKey = "Settings",  #Head Keyword in the xml
    $ParentWindow = $null
  )

  $xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   MinWidth="400"
   Width="400"
   MinHeight ="400"
   SizeToContent="Height"
   Title="Config Dialog"
   Topmost="True">

   
   <Grid Margin="10,10,10,10">
      <Grid.ColumnDefinitions>
         <ColumnDefinition Width="Auto"/>
         <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
        <TextBlock Grid.Column="0" Grid.Row="0" Grid.ColumnSpan="2" Margin="5"  FontWeight="Bold">Your Settings</TextBlock>
        <StackPanel Name ="StackKeys" Grid.Column="0" Grid.Row="1">
            <!-- TextBlock Margin="5">Name</>
            <TextBlock Margin="5">Email</TextBlock -->
        </StackPanel>
        <StackPanel Name ="StackValues"  Grid.Column="1" Grid.Row="1">
            <!-- TextBox Margin="5"></>
            <TextBox Margin="5"></TextBox -->
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,10,0,0" Grid.Row="3" Grid.ColumnSpan="2">
        <Button Name="ButOk" MinWidth="80" Height="22" Margin="5">OK</Button>
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

  $window.ButOk.add_Click(
    {
      $window.DialogResult = $true
    }
  )
  #endregion Event Handlers

  #region Manipulate Window Content
  $window.Title = $Title
  $Margin = New-Object System.Windows.Thickness -ArgumentList @("5")
  foreach($key in $SettingsXml.$EntryKey.ChildNodes ){
  
    $KeyBlock = New-Object System.Windows.Controls.TextBlock
    #$KeyBlock.Name = $key.Name
    $KeyBlock.Text = $key.Name
    $keyBlock.Margin = $Margin
    $keyBlock.Height = 20
    $window.StackKeys.AddChild($KeyBlock)
    
    
    #Values
    if($key.Type -eq "switch")
    {
      $WPFElement = New-Object System.Windows.Controls.CheckBox
      $WPFElement.IsChecked = $key.InnerText -eq "true"

    } elseif ($key.Type -eq "Password")
    {
      $WPFElement= New-Object System.Windows.Controls.PasswordBox
      $WPFElement.Password = $key.InnerText

    } else 
    {
      $WPFElement = New-Object System.Windows.Controls.TextBox
      $WPFElement.Text = $key.InnerText
    }
    
    $WPFElement.Height = 20
    $WPFElement.Name = $key.Name
    $WPFElement.Margin = $Margin
    $window.StackValues.AddChild($WPFElement)
    
    
  }
  
  #Center to Parent
  if($ParentWindow -ne $null){
    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        
    $Window.Owner = $ParentWindow
  } else {
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
  }

  #endregion

  # Show Window
  $result = Show-WPFWindow -Window $window

  #process the result
  if($result -eq "ok"){
    foreach($key in $window.StackValues.Children){
  
      $tok = $key.Name


      #Values
      if($SettingsXml.$EntryKey.$tok.type -eq "Switch")
      {
        $SettingsXml.$EntryKey[$tok].innerText = $key.IsChecked

      } elseif ($SettingsXml.$EntryKey.$tok.type -eq "Password")
      {
    
        $SettingsXml.$EntryKey[$tok].innerText = $Key.Password

      } else 
      {
        $SettingsXml.$EntryKey[$tok].innerText = $Key.Text
      }
    
    }
  }
  
  return  $SettingsXml
}

