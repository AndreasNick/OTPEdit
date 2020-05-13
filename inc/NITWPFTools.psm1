$Global:Shell32Symbols = @{"GearW10Dll"=316;"Star"=43;"Star2"=208;"Warning"=235;"X"=131;"Check"=144;"CheckGreen"=302;"Users"=160;"Key"=44;"UsersAndKey"=111;"Refresh"=238;"Gear"=90;"Error"=109;"Information"=277}

function Get-TwitterBitmapBase64{
  #Base64 encoded png
  $bm=@'
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
'@

  return $bm
}

function Get-IconFromShell32dll
{
  param($index)

  #https://social.technet.microsoft.com/Forums/windows/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27/using-icons-in-powershell-scripts?forum=winserverpowershell


  $code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@

  Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing
  Return [System.IconExtractor]::Extract("shell32.dll", $index, $true)
}


function Get-ImageSourceFromShell32dll
{
  param 
  (
    $IconIndex = 1
  )
  Add-Type -AssemblyName WindowsBase
  $Icon = Get-IconFromShell32dll -index $IconIndex
  $source = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($icon.Handle, [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
  return $source
}

function Convert-ImageToBase64{
  param($Image)
  Add-Type -AssemblyName System.Drawing
  $Bitmap = new-Object System.Drawing.Bitmap($image.ToBitmap(), 128,128)

  $ms = new-object System.IO.MemoryStream
  $Bitmap.save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  $ImageBase64 = [Convert]::ToBase64String($ms.GetBuffer())
  $ms.Dispose()
  $Bitmap.Dispose()
  
  return $ImageBase64
}

function Convert-BitmapToBae64{

  param($Bitmap)
  Add-Type -AssemblyName System.Drawing
  $ms = new-object System.IO.MemoryStream
  $Bitmap.save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  $ImageBase64 = [Convert]::ToBase64String($ms.GetBuffer())
  $ms.Dispose()
  $Bitmap.Dispose()

  return $ImageBase64
}


function Convert-Base64ToImageSourtce{
 
  param($bgImage64)
 
  Add-Type -AssemblyName PresentationCore
  
  [byte[]] $binaryData = [System.Convert]::FromBase64String($bgImage64)

  $bi = new-object System.Windows.Media.Imaging.BitmapImage
  $bi.BeginInit()
  $bi.StreamSource = new-Object  System.IO.MemoryStream(@(,$binaryData))
  $bi.EndInit()
  
  return $bi
}




function New-MessageBox{
  <#
      .SYNOPSIS
      Open a Message Dialog

      .DESCRIPTION
      Open a Message Dialog

      .PARAMETER Title
      Title of the Window

      .PARAMETER Symbol
      Describe parameter -Symbol.

      .PARAMETER Message
      Describe parameter -Message.

      .PARAMETER Hyperlinks
      Describe parameter -Hyperlinks.

      .PARAMETER linksAltText
      Describe parameter -linksAltText.

      .PARAMETER EndMessage
      Describe parameter -EndMessage.

      .PARAMETER DisableCancle
      Describe parameter -DisableCancle.

      .PARAMETER ParentWindow
      Describe parameter -ParentWindow.

      .EXAMPLE
      About Dialog
      New-MessageBox  -Symbol $Global:Shell32Symbols.Star -title "About" -Message $("Andreas Nick - 2020`n") 
      -Hyperlinks @("https://software-Virtualisierung.de", "https://andreasnick.com","https://nick-it.de") 
      -linksAltText @("My German blog Software Sofrtualisierung`n", "My English blog`n","My Company`n") 
      -EndMessage "`nThanks to thorsten for the idea and the testing`n Contact: info@nick-it.de" -DisableCancle $true

      .NOTES

      .LINK

      .INPUTS

      .OUTPUTS
  #>


  param
  (
    [String] $Title = "New Message",
    [ArgumentCompleter({ @(,$Global:Shell32Symbols.keys)  })]
    [String] $Symbol = "235", #Warning,
    [String] $Message = "Empty",
    [String[]] $Hyperlinks = @(),
    [String[]] $linksAltText = @(), #Visible Text for Hyperlins
    [String] $EndMessage = $null,
    [Bool] $DisableCancle = $false,
    [System.Windows.Window] $ParentWindow

  )
  $xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   MinWidth="200"
   Width ="500"
   SizeToContent="Height"
   Title="New Message"
   Topmost="True" ResizeMode="NoResize">
   <Grid Margin="10,20,10,10">
      <Grid.ColumnDefinitions>
         <ColumnDefinition Width="Auto"/>
         <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

        <Image Name="ImageSymbol" Width="64" Height="64" Grid.Row="1" Grid.Column="0" Margin="10"></Image>
        <TextBlock Name="TextBlockMessage" TextWrapping="Wrap"  Grid.Row="1" Grid.Column="1" Margin="10">
            
            <!-- Hyperlink Name="Hyperlink1" NavigateUri="https://andreasnick.com">some site</-->
            
        </TextBlock>
      <StackPanel Name = "ButtonStack" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="0,10,0,0" Grid.Row="2" Grid.ColumnSpan="2">
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

  $window.ImageSymbol.Source = Get-ImageSourceFromShell32dll -IconIndex $Symbol
  $window.TextBlockMessage.AddText($Message)
  
  if($Hyperlinks.count -ne 0){
    for($i=0; $i -le $Hyperlinks.count;$i++){
      $HypLnk = New-Object System.Windows.Documents.Hyperlink
      
      $HypLnk.NavigateUri = $Hyperlinks[$i]
      
      $HypLnk.AddText($linksAltText[$i])
      
      $HypLnk.Name = "HypLink_$i"
      $HypLnk.add_Click({
          param
          (
            [Parameter(Mandatory)][Object]$sender,
            [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
          )
          
          #Start-Process $e.OriginalSource.NavigateUri
          New-PSDrive -Name HKCR -PSProvider registry -Root Hkey_Classes_Root | Out-Null
          $browserPath = ((Get-ItemProperty 'HKCR:\http\shell\open\command').'(default)').Split('"')[1]
          & $browserPath $e.OriginalSource.NavigateUri
        }
      )
      $window.TextBlockMessage.AddChild($HypLnk)
    }
  }
  
  if($EndMessage -ne $null){
    $window.TextBlockMessage.AddText($EndMessage)
  }  
  
  if($Title -ne $null){
    $window.Title = $Title
  }
  
  if($DisableCancle){
    $Window.ButtonStack.children.remove($Window.ButCancel)
  }
  
  #Center to Parent
  if($ParentWindow -ne $null){
    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $Window.Owner = $ParentWindow
  } else {
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
  }
  
  
  $Window.Icon = Get-ImageSourceFromShell32dll -IconIndex $Symbol
  #endregion

  # Show Window
  $result = Show-WPFWindow -Window $window
  
  Return $Result
  
  #endregion Process results
}


function Out-Message{

  [CmdletBinding()]
  param(
    [String] $Message = "null",
    [switch] $UseMsgBox = $true,
    [ValidateSet("Error", "Warning", "Information")]
    [String] $Type = "Information",
    [bool] $DisableCancel = $true,
    $ParentWindow = $null

        
  )



  if($UseMsgBox) {
    Return New-MessageBox -Title $Type -Symbol $Global:Shell32Symbols.$Type  -Message $Message -ParentWindow $ParentWindow -DisableCancle $DisableCancel
    #return [System.Windows.Forms.MessageBox]::Show("$Message","Error",[System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::$Type)
    
    
  } else {
    Write-Verbose $Message -Verbose
  }
}
