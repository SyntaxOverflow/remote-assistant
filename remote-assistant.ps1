Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Get control over the console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

#---GUI---
$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = "300,120"
$Form.Text = "Remote Assistant"
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox = $false
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $True
$Form.Add_KeyDown({if($_.KeyCode -eq 'Enter'){$Button.PerformClick()}})

$Input = New-Object System.Windows.Forms.ComboBox
$Input.Location = New-Object System.Drawing.Point(5,5)
$Input.Size = New-Object System.Drawing.Size(290,30)
$Input.Font = "Microsoft Sans Serif,12"
$Input.Text = ""
$Input.MaxDropDownItems = "25"

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Point(5,40)
$Button.Size = New-Object System.Drawing.Size(290,30)
$Button.Font = "Microsoft Sans Serif,12"
$Button.Text = "Engage"
$Button.Add_Click({letsgo})

$Status = New-Object System.Windows.Forms.Label
$Status.Location = New-Object System.Drawing.Point(5,75)
$Status.Size = New-Object System.Drawing.Size(290,40)
$Status.TextAlign = "MiddleCenter"
$Status.Font = "Microsoft Sans Serif,12"
$Status.Text = "Ready"

$Form.Controls.AddRange(@($Input,$Button,$Status))

#---function---
function letsgo{
    $Search = $Input.Text
    $Status.Text = "Searching for PC"
    $PC = Get-ADComputer -Filter "Name -like '*$Search*'" | Select-Object -ExpandProperty Name

    if($PC.Count -gt 1){
        $Status.Text = "More than one PC found"
        $PC | foreach{$Input.Items.Add($_)}
        $Input.Text = "Please select PC"
        return
    }elseif($PC -eq $null){
        $Status.Text = "No PC found. Connect anyway. (IP-Address?)"
        msra.exe /offerra $Search
        return
    }

    $Input.Text = "$PC"
    $Status.Text = "Testing Connection"
    if(Test-NetConnection -ComputerName $PC -InformationLevel Quiet){
        $Status.Text = "Connecting..."
        msra.exe /offerra $PC
        $Input.Items.Clear()
        $Input.Text =""
        $Status.Text = "Ready"
    }else{
        $Status.Text = "$PC not reachable!"
    }
}

#Hide Console when GUI appears and bring it back if GUI gets closed
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
$Form.ShowDialog()
[Console.Window]::ShowWindow($consolePtr, 4)
