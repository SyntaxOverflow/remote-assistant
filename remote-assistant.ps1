#https://github.com/HeiligerMax
#Licensed under GNU General Public License v3.0

#Get control over the Powershell window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
###
#Add WPF
Add-Type -AssemblyName PresentationFramework
[xml]$xaml = @"
<Window x:Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Remote-Assistent" Width="300" Height="200" WindowStartupLocation="CenterScreen" FontFamily="Microsoft Sans Serif" FontSize="16" SizeToContent="Height" ResizeMode="CanMinimize" MinWidth="300" MinHeight="100">
    <Grid>
        <ComboBox x:Name="InputField" Margin="10,10,10,0" Padding="5" IsEditable="True" VerticalAlignment="Top" ToolTip="The whole or part of Computername"/>
        <Button x:Name="Button" Content="Search and connect" Margin="10,50,10,0" Padding="5" VerticalAlignment="Top" IsDefault="True" ToolTip="Search for the Computername in AD"/>
        <TextBox x:Name="OutputField" Margin="10,90,10,10" Padding="5" TextWrapping="Wrap" Text="Ready" HorizontalContentAlignment="Center" IsReadOnly="True" ToolTip="https://github.com/HeiligerMax"/>
    </Grid>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
###
$InputField = $window.FindName("InputField")
$Button = $window.FindName("Button")
$OutputField = $window.FindName("OutputField")
$Button.Add_Click{(offerra)}
###
function offerra{
    $Search = $InputField.Text
    $OutputField.Text = "Suche nach APC"
    $PC = Get-ADComputer -Filter "Name -like '*$Search*'" | Select-Object -ExpandProperty Name

    if($PC.Count -gt 1){
        $OutputField.Text += "`r`nMore than one Computer found"
        $PC | ForEach-Object{$InputField.Items.Add($_)}
        $InputField.Text = "Please select Computer"
        $InputField.IsDropDownOpen = $true
        return
    }elseif($null -eq $PC){
        $OutputField.Text += "`r`nNo Computer found."
        $BoxAntwort = [System.Windows.MessageBox]::Show("No Computer found!`r`nConnect to $Search anyway?",":(","YesNo","Error")
        if($BoxAntwort -eq "Yes"){
            $OutputField.Text += "`r`Connecting"
            msra.exe /offerra $Search
        }
        $OutputField.Text = "Ready"
        return
    }

    $InputField.Text = "$PC"
    $OutputField.Text += "`r`nTesting connection"
    if(Test-NetConnection -ComputerName $PC -InformationLevel Quiet){
        $OutputField.Text += "`r`Connecting`r`n"
        msra.exe /offerra $PC
        $InputField.Items.Clear()
        $InputField.Text =""
        $OutputField.Text = "Ready"
    }else{
        $OutputField.Text += "`r`n$PC not reachable!"
    }
}
###
#Hide Powershell when GUI appears and bring it back if GUI gets closed
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
$window.ShowDialog() | Out-Null
[Console.Window]::ShowWindow($consolePtr, 4)
