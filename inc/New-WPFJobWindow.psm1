

function New-RunspacePool {

  param (   
    [Int][ValidateRange(2,200)] $ThrottleLimit = 32,
    [Switch] $UseLocalVariables
  )
  
  # Initialisierungsarbeiten durchfï¿½hren
  # einen initialen Standard-Zustand der PowerShell beschaffen:

  $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
  # darin auf Wunsch alle lokalen Variablen einblenden:
  if ($UseLocalVariables)
  {
    # zuerst in einer "frischen" PowerShell alle Standardvariablen ermitteln:
    $ps = [Powershell]::Create()
    $null = $ps.AddCommand('Get-Variable')
    $oldVars = $ps.Invoke().Name
    $ps.Runspace.Close()
    $ps.Dispose()

    # nun aus der vorhandenen PowerShell alle eigenen Variablen in der neuen PowerShell einblenden,
    # (die nicht zu den Standardvariablen zählen):
    Get-Variable | 
    Where-Object { $_.Name -notin $oldVars } |
    Foreach-Object {
      $SessionState.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($_.Name, $_.Value, $null)))
    }
  }

  # einen Runspace-Pool mit den nätigen Threads anlegen:
  $RunspacePool = [Runspacefactory]::CreateRunspacePool(1, $ThrottleLimit, $SessionState, $host)
  $RunspacePool.ApartmentState = "STA"
  $RunspacePool.ThreadOptions = "ReuseThread"
  $RunspacePool.Open() 

  # in dieser Liste die aktuell noch laufenden Threads vermerken:
  $ThreadList = New-Object System.Collections.ArrayList        
  $RunspacePool  | Add-Member -MemberType NoteProperty -Name ThreadList -Value $ThreadList
  $RunspacePool  | Add-Member -MemberType NoteProperty -Name threadID -value 0
  
  return $RunspacePool
}

function New-RunspaceJob{
  param(
    [Parameter(Mandatory=$true)][Management.Automation.Runspaces.RunspacePool] $RunspacePool,
    [Parameter(Mandatory=$true)] $Code,
    $Arguments,
    
    [Int] $TimeoutSec = -1
    
  )
  # Thread anlegen:
  $powershell = [powershell]::Create()
  
  
  #$null = $PowerShell.AddScript($Code).AddArgument($Arguments).
  $env = $PowerShell.AddScript($Code)
  Foreach($item in $Arguments) {
    #Write-Verbose "Add Argument : $item" 
    $env.AddArgument($item)
  }
  
  $powershell.RunspacePool = $RunspacePool
  # Informationen ï¿½ber diesen Thread in einem Hashtable speichern:
  $RunspacePool.threadID++
  
  Write-Verbose "Starte Thread $($RunspacePool.threadID)"

  $threadInfo = @{
    PowerShell = $powershell
    StartTime = Get-Date
    ThreadID = $RunspacePool.threadID
    Runspace = $powershell.BeginInvoke()
    TimeoutSec = $TimeoutSec
  }

  # diese Information in der Liste der laufenden Threads vermerken:
  $null = ($RunspacePool.ThreadList).Add($threadInfo)
  
  return $threadInfo
}

function Get-Runspaces{
  
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][Management.Automation.Runspaces.RunspacePool] $RunspacePool
  )
    
  $aborted = 0 
  #try
  #{

  if($RunspacePool.ThreadList.Count -ne 0){
    Write-Verbose $("Threads:" +$RunspacePool.ThreadList.Count)
  }

  # alle noch vorhandenen Threads untersuchen:
  Write-Verbose $("Threads:" +$RunspacePool.ThreadList.Count)
    
  Foreach($thread in $RunspacePool.ThreadList) {
    Write-Verbose "Thread $($thread.ThreadID)"

    If ($thread.Runspace.isCompleted) {
      # wenn der Thread abgeschlossen ist, Ergebnis abrufen und
      # Thread als "erledigt" kennzeichnen:
      if($thread.powershell.Streams.Error.Count -gt 0) 
      {
        # falls es zu Fehlern kam, Fehler ausgeben:
        foreach($ErrorRecord in $thread.powershell.Streams.Error) {
          Write-Error -ErrorRecord $ErrorRecord
        }
      }
      if ($thread.TimedOut -ne $true)
      {
        # Ergebnisse des Threads lesen:
        $thread.powershell.EndInvoke($thread.Runspace)
        Write-Verbose "empfange Thread $($thread.ThreadID)"
      }
      $thread.Done = $true
    }
    # falls eine maximale Laufzeit festgelegt ist, diese ï¿½berprï¿½fen:
    elseif ($thread.TimeoutSec -gt 0 -and $thread.TimedOut -ne $true)
    {
      # Thread abbrechen, falls er zu lange lief:
      $runtimeSeconds = ((Get-Date) - $thread.StartTime).TotalSeconds
      if ($runtimeSeconds -gt $thread.TimeoutSec)
      {
        Write-Warning -Message "Thread $($thread.ThreadID) timed out."
        $thread.TimedOut = $true
        $null = $thread.PowerShell.BeginStop({}, $null)
      }
    }
  }

  # alle abgeschlossenen Threads ermitteln:
  $ThreadCompletedList = $RunspacePool.ThreadList | Where-Object { $_.Done -eq $true }
  if ($ThreadCompletedList.Count -gt 0)
  {
    # diese Threads aus der Liste der aktuellen Threads entfernen:
    foreach($threadCompleted in $ThreadCompletedList)
    {
      Write-Verbose "Thread compleaded $($thread.ThreadID)"
      # Thread entsorgen:
      $threadCompleted.powershell.Stop()
      $threadCompleted.powershell.dispose()
      $threadCompleted.Runspace = $null
      $threadCompleted.powershell = $null
      $RunspacePool.ThreadList.remove($threadCompleted)
    }
  }
}

function Clear-Runspace{
  param(
    [Parameter(Mandatory=$true)][Management.Automation.Runspaces.RunspacePool] $RunspacePool
  )
    
  # falls es noch laufende Threads gibt (Benutzer hat CTRL+C gedrï¿½ckt)
  # diese abbrechen und entsorgen:
  
  foreach($thread in $RunspacePool.ThreadList)
  {
    $thread.powershell.dispose() 
    $thread.Runspace = $null
    $thread.powershell = $null
  }
  # RunspacePool schließen:
  $RunspacePool.close()
  # Speicher aufräuumen:

  [GC]::Collect() 
}

function Remove-Runspace{
  param(
    [Parameter(Mandatory=$true)][Management.Automation.Runspaces.RunspacePool] $RunspacePool,
    [int] $ThreadID
  )
    
  # falls es noch laufende Threads gibt (Benutzer hat CTRL+C gedrï¿½ckt)
  # diese abbrechen und entsorgen:
  
  foreach($thread in $RunspacePool.ThreadList)
  {
    if($ThreadID -eq $thread.ThreadId){
      Write-Verbose "Kill Runspace $ThreadID" 
      $thread.TimeoutSec=$true
    }
  }
  [GC]::Collect() 
}


function New-VisualJobWindow{
  param
  (
    $ParentWindow = $null, #Center Parent
    [string] $Title = "Please Wait",
    [string] $ActionLabel = "In Action"

  )



    
  #region XAML window definition
  # Right-click XAML and choose WPF/Edit... to edit WPF Design
  # in your favorite WPF editing tool
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName WindowsFormsIntegration
  $xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   Height="150"
   Width ="400"
   SizeToContent="Height"
   Title="Launcher"
    WindowStyle="ToolWindow"
    Background="Azure"  
    ResizeMode="NoResize"
    
   Topmost="True">
   <Window.Resources>
   
    <Color x:Key="FilledColor" A="255" B="128" R="0" G="38"/>
    <Color x:Key="UnfilledColor" A="0" B="128" R="0" G="38"/>

    <Style x:Key="BusyAnimationStyle" TargetType="Button">
        <Setter Property="Background" Value="Transparent"/>

        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Control">
                    <ControlTemplate.Resources>
                        <Storyboard x:Key="Animation0" BeginTime="00:00:00.0" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse0" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation1" BeginTime="00:00:00.1" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse1" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation2" BeginTime="00:00:00.2" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse2" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation3" BeginTime="00:00:00.3" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse3" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation4" BeginTime="00:00:00.4" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse4" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation5" BeginTime="00:00:00.5" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse5" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation6" BeginTime="00:00:00.6" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse6" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>

                        <Storyboard x:Key="Animation7" BeginTime="00:00:00.7" RepeatBehavior="Forever">
                            <ColorAnimationUsingKeyFrames Storyboard.TargetName="ellipse7" Storyboard.TargetProperty="(Shape.Fill).(SolidColorBrush.Color)">
                                <SplineColorKeyFrame KeyTime="00:00:00.0" Value="{StaticResource FilledColor}"/>
                                    <SplineColorKeyFrame KeyTime="00:00:00.8" Value="{StaticResource UnfilledColor}"/>
                            </ColorAnimationUsingKeyFrames>
                        </Storyboard>
                    </ControlTemplate.Resources>

                    <ControlTemplate.Triggers>
                        <Trigger Property="IsVisible" Value="True">
                            <Trigger.EnterActions>
                                <BeginStoryboard Storyboard="{StaticResource Animation0}" x:Name="Storyboard0" />
                                <BeginStoryboard Storyboard="{StaticResource Animation1}" x:Name="Storyboard1"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation2}" x:Name="Storyboard2"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation3}" x:Name="Storyboard3"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation4}" x:Name="Storyboard4"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation5}" x:Name="Storyboard5"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation6}" x:Name="Storyboard6"/>
                                <BeginStoryboard Storyboard="{StaticResource Animation7}" x:Name="Storyboard7"/>
                            </Trigger.EnterActions>

                            <Trigger.ExitActions>
                                <StopStoryboard BeginStoryboardName="Storyboard0"/>
                                <StopStoryboard BeginStoryboardName="Storyboard1"/>
                                <StopStoryboard BeginStoryboardName="Storyboard2"/>
                                <StopStoryboard BeginStoryboardName="Storyboard3"/>
                                <StopStoryboard BeginStoryboardName="Storyboard4"/>
                                <StopStoryboard BeginStoryboardName="Storyboard5"/>
                                <StopStoryboard BeginStoryboardName="Storyboard6"/>
                                <StopStoryboard BeginStoryboardName="Storyboard7"/>
                            </Trigger.ExitActions>
                        </Trigger>
                    </ControlTemplate.Triggers>

                    <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}">
                        <Grid>
                        <Canvas Height="60" Width="60">
                            <Canvas.Resources>
                                <Style TargetType="Ellipse">
                                    <Setter Property="Width" Value="8"/>
                                    <Setter Property="Height" Value="8" />
                                    <Setter Property="Fill" Value="#009B9B9B" />
                                </Style>
                            </Canvas.Resources>

                            <Ellipse x:Name="ellipse0" Canvas.Left="1.75" Canvas.Top="21"/>
                            <Ellipse x:Name="ellipse1" Canvas.Top="7" Canvas.Left="6.5"/>
                            <Ellipse x:Name="ellipse2" Canvas.Left="20.5" Canvas.Top="0.75"/>
                            <Ellipse x:Name="ellipse3" Canvas.Left="34.75" Canvas.Top="6.75"/>
                            <Ellipse x:Name="ellipse4" Canvas.Left="40.5" Canvas.Top="20.75" />
                            <Ellipse x:Name="ellipse5" Canvas.Left="34.75" Canvas.Top="34.5"/>
                            <Ellipse x:Name="ellipse6" Canvas.Left="20.75" Canvas.Top="39.75"/>
                            <Ellipse x:Name="ellipse7" Canvas.Top="34.25" Canvas.Left="7" />
                            <Ellipse Width="39.5" Height="39.5" Canvas.Left="8.75" Canvas.Top="8" Visibility="Hidden"/>
                        </Canvas>
                            <Label Content="{Binding Path=Text}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Grid>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    </Window.Resources>

   <Grid Margin="10,10,10,10">
        <StackPanel Orientation="Vertical">
        <Label Name="Label1" >Launching App-V application</Label>
            <Button Style="{StaticResource BusyAnimationStyle}"  BorderBrush="Transparent" BorderThickness="0">test</Button>
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

  #endregion Code Behind

  #region Convert XAML to Window
  $window = Convert-XAMLtoWindow -XAML $xaml 
  #endregion

  #region Define Event Handlers
  # Right-Click XAML Text and choose WPF/Attach Events to
  # add more handlers

  #endregion Event Handlers
  $window.Add_Closing({
      #[System.Windows.Forms.Application]::Exit()
      # Stop-Process $pid
      
      
  })
  
  
  
  
  
  #region Manipulate Window Content

  #Center to Parent
  if($ParentWindow -ne $null){
    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        
    $Window.Owner = $ParentWindow
  } else {
    $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
  }
  
  $window.Title = $Title
  $window.Label1.Content = $ActionLabel

  #endregion
  
  Return $window
}


function Show-VisualJobWindow{
  [CmdletBinding()]
  Param($window)
 
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
  
  $result = Show-WPFWindow -Window $window
  Write-Verbose "Close JobWindow"

}


<# Example

    $win = new-VisualJobWindow

    $pool = New-RunspacePool -UseLocalVariables

    $job = New-RunspaceJob -RunspacePool $pool -Arguments $win -Code {
    param($window)
    Get-ChildItem "C:\Windows"   | out-File "C:\Profiles\Administrator\Desktop\out.txt" 
  
    Start-Sleep -Seconds 5
    Write-Verbose "Finish Job" -Verbose
  
    #Close Window
    $window.Dispatcher.InvokeAsync{$result = $window.close()}.Wait()
    #$window | Out-String | out-File "C:\Profiles\Administrator\Desktop\out.txt" -Append
  
    Write-Verbose "Finished" -Verbose
    } -Verbose

    # Show Window

    Show-VisualJobWindow -window $win
    Clear-Runspace -RunspacePool $pool

#>
