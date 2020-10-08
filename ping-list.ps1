Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -assembly System.Windows.Forms

#title for the winform
$Title = "Ping Progress:"
#winform dimensions
$height=100
$width=400
#winform background color
$color = "White"

#create the form
$form1 = New-Object System.Windows.Forms.Form
$form1.Text = $title
$form1.Height = $height
$form1.Width = $width
$form1.BackColor = $color
$form1.ShowIcon = $false


$form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
#display center screen
$form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# create label
$label1 = New-Object system.Windows.Forms.Label
$label1.Text = "not started"
$label1.Left=5
$label1.Top= 10
$label1.Width= $width - 20
#adjusted height to accommodate progress bar
$label1.Height=15
$label1.Font= "Verdana"
#optional to show border 
#$label1.BorderStyle=1

#add the label to the form
$form1.controls.add($label1)

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Step = 1
$progressBar1.Style="Continuous"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = $width - 40
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size

$progressBar1.Left = 15
$progressBar1.Top = 40

$form1.Controls.Add($progressBar1)
$label1.text="Preparing to start"
$form1.Refresh()
start-sleep -Seconds 1

$timer = New-Object System.Windows.Forms.Timer 
$timer.Interval = 1000

function script:startProcess{

if($MyInvocation.MyCommand.CommandType -eq "ExternalScript"){
$ScriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition
}else{
$ScriptPath = Split-Path -Path ([Environment]::GetCommandLineArgs()[0])
if(!$ScriptPath){$ScriptPath = "."}}

#if you want to use with powershell just change dir to
#$dir = $PSScriptRoot
$dir = $ScriptPath
if((Test-Path -Path "$dir\test.txt")){
    $pcIds = Get-Content("$dir\test.txt")
    if($pcIds.Length -gt 0){
    [int] $count = 0
    [int] $activeCount = 0
    [int] $inactiveCount = 0
    $progressBar1.Maximum = $pcIds.Length

    foreach($pcId in $pcIds){
        if(Test-Connection -ComputerName $pcId -Count 1 -ErrorAction SilentlyContinue){
            $success += $pcId + "`n"
            $activeCount++
        }else{
            $failure += $pcId + "`n"
            $inactiveCount++
        }
        $count++
        $pct = ($count/$pcIds.Length) * 100
        $pct = [Math]::Round($pct)
        $progressBar1.Value = $count
        $label1.text="Progress: $pct %"
        $form1.Text = "Progressing: $count/" +$pcIds.Length  
        $form1.Refresh()
    }
    $form1.Refresh()
    $timer.Enabled = $false
    if (!(Test-Path -Path "$dir\succeed.txt"))
    {
        New-Item -path $dir -name succeed.txt -type "file"
    }
    if(!(Test-Path -Path "$dir\failed.txt")){
        New-Item -path $dir -name failed.txt -type "file"
    }
    $success | Out-File "$dir\succeed.txt"
    $failure | Out-File "$dir\failed.txt"
    
    $messageText = "Files saved in script folder.`nActive pc count is $activeCount`nInactive pc count is $inactiveCount"
    $messageButtons = New-Object System.Windows.MessageBoxButton::OK
    $messageIcon = [System.Windows.MessageBoxImage]::Information
    $messageTitle = "Search completed"
    $messageResult = [System.Windows.MessageBox]::Show($messageText,$messageTitle,$messageButtons,$messageIcon) 
    if($messageResult -eq [System.Windows.MessageBoxResult]::OK){
        $timer.Enabled = $false
        $timer.Stop()
        $form1.Dispose()
        $timer.Dispose()
        Start-Sleep -s 5
        return
    }
   }else{
    $form1.Close()
    $timer.Enabled = $false
    $timer.Stop()
    $form1.Dispose()
    $timer.Dispose()
    $errorMessageBody = "test.txt file doesn't have any content.`nThe file needs 1 item at least. "
    $errorTitle = "Error"
    $errorButton = [System.Windows.MessageBoxButton]::OK
    $errorMessage = [System.Windows.MessageBox]::Show($errorMessageBody, $errorTitle, $errorButton, [System.Windows.MessageBoxImage]::Error)
    return
   }
  }else{
    $form1.Close()
    $timer.Enabled = $false
    $timer.Stop()
    $form1.Dispose()
    $timer.Dispose()
    $errorMessageBody = "test.txt file does not exist in script folder."
    $errorTitle = "Error"
    $errorButton = [System.Windows.MessageBoxButton]::OK
    $errorMessage = [System.Windows.MessageBox]::Show($errorMessageBody, $errorTitle, $errorButton, [System.Windows.MessageBoxImage]::Error)
    return
  }
}


$timer.add_Tick({
    startProcess
})

if($progressBar1.Value -eq 0){
$timer.Enabled = $true
$timer.Start()

$form1.Add_Shown({$form1.Activate()})
$form1.ShowDialog()
}


