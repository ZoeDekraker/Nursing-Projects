# Program takes a csv file of therapy entries, groups by week of year and then by entries per week, 
# displaying result on Out-Gridview. Relies on "Date" column header and a date format of dd/mm/yyyy. 


Add-Type -AssemblyName PresentationFramework

# function to get frequency per week in each week of year.
Function Get-PMP()
{
    Import-Csv $Global:file_path |
    Select-Object *, @{Name="Week_Number";Expression={Get-Date -UFormat %V $_.Date}} |
    Group-Object -Property Week_Number | Add-Member -MemberType AliasProperty -Name WeekOfYear -Value Name -PassThru |
    Add-Member -MemberType AliasProperty -Name TotalEntries -Value Count -PassThru | Select-Object -Property WeekOfYear, TotalEntries | 
    Out-GridView -Title 'PMP Audit Results'
}

# function to select csv file to audit
Function Get-FileName()
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “CSV files (*.csv)| *.csv”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Filename

}

# GUI XAML file (WPF GUI)
$xamlFile = "MainWindow.xaml"
$inputXAML = Get-Content -Path $xamlFile -Raw

# Tidy up XAML to work with PowerShell
$inputXAML = $inputXAML -replace 'mc:Ignorable="d"', '' -replace "x:N", "N" -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXAML

# read XAML in
$reader = New-Object System.Xml.XmlNodeReader $XAML

# load & catch the error if it doesn't load.
try {
    $psform = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Output $_.Exception
    throw
}

# create controls already named in XAML
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $psform.FindName($_.Name) -ErrorAction Stop
    }
    catch {
        throw
    }
}

# select_button action - get filename and update label
$var_select_file_button.Add_Click({$Global:file_path = Get-FileName
    $var_file_label.Content = $Global:file_path  
})

# start button action - analyse file and output to gridview.
$var_start_button.Add_Click({Get-PMP})

# show the form
$psform.ShowDialog()
