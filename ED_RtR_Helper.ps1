<################################################################
| Name: ED_RtR_Helper.ps1                                      |
| Author: Alex Christensen                                     |
| Date: 06/28/2017                                             |
| Purpose: Keeps track of systems that have been visited along |
|          the exploration road to riches.                     |
| Usage: Run in powershell.                                    |
|                                                              |
|--------------------------------------------------------------|
| Update History                                               |
|                                                              |
| 06/28/2017 CHRIALE: Created Script.                          |
|                     Database loading completed.              |
| 06/29/2017 CHRIALE: Added user settings. started UI.         |
| 06/30/2017 CHRIALE: Finished UI. Dataobjects not persisting. |
| 07/01/2017 CHRIALE: Dataobjects fixed. Find/Replace function |
|                     not replacing.                           |
| 07/01/2017 CHRIALE: Script is Completed.                     |
|                                                              |
################################################################>
 
#region /* Load External References */
 
# Adds the abilily to hide the console window
$MemDef = @"
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
[DllImport("user32.dll")]
public static extern void DisableProcessWindowsGhosting();
"@

Add-Type -Name Window -Namespace Console -MemberDefinition $MemDef

 
# imports the Windows Forms and Drawing Functions
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
 
 
# Called function to hide the console
function HideConsole(){
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
 
    #also disables the Not responding timeout
    [Console.Window]::DisableProcessWindowsGhosting()
}
 
#endregion /* Load External References */
 
#region /* Script Data Objects */
 
# declares a new system object.
function NewSystemObject()
{

    $emptySystemObject = [PSCustomObject]@{
        Name = ""
        X = ""
        Y = ""
        Z = ""
        ID = ""
        Visited = $false
        Planets = @()
    }
 
    return $emptySystemObject
}
 
# declares a new planet object.
function NewPlanetObject()
{
    $emptyPlanetObject= [PSCustomObject]@{
        Name = ""
        Distance = ""
        Type = ""
        ID = ""
        Visited = $false
    }
    
    return $emptyPlanetObject
}

# Object to hold the current user settings
function NewUserSettingsObject()
{
    $emptyUserSettingsObject = [PSCustomObject]@{
        AutomarkSystem = $false
        AutomarkPlanets = $false
        SkipVisited = $false
        CurrentSystem = NewSystemObject
        CurrentRoute = @()
    }

    $emptyUserSettingsObject.CurrentSystem.Name = "No System Loaded"

    return $emptyUserSettingsObject
}
 
#endregion  /* Script Data Objects */
 
#region /* Global Form Objects */
 
# the windows forms objects
$mainForm = New-Object System.Windows.Forms.Form
$systemLabel = New-Object System.Windows.Forms.Label
$systemVisitedCheckbox = New-Object System.Windows.Forms.CheckBox
$planetListDataGridView = New-Object System.Windows.Forms.DataGridView
$importTextBox = New-Object System.Windows.Forms.TextBox
$importButton = New-Object System.Windows.Forms.Button
$automarkStarCheckBox = New-Object System.Windows.Forms.CheckBox
$automarkPlanetsCheckBox = New-Object System.Windows.Forms.CheckBox
$skipVisitedCheckBox = New-Object System.Windows.Forms.CheckBox
$nextButton = New-Object System.Windows.Forms.Button
$numberLeftLabel = New-Object System.Windows.Forms.Label
$exitButton = New-Object System.Windows.Forms.Button
$labelFont = New-Object System.Drawing.Font("MS Sans Serif", 10)
$labelFontBig = New-Object System.Drawing.Font("MS Sans Serif", 14)

# the default icon path
$iconPath = "$($PSScriptRoot)\Data\Icon.ico"

# The main Data Objects for the script
$mainSystemDatabaseObject | Out-Null

# The user settings object
$mainUserSettingsObject | Out-Null
 
#endregion /* Global Form Objects */
 
#region /* XML Communicator Functions */

# returns the Type Translation for initial db build
function TranslateTypeCode($code)
{
    switch($code)
    {
        26 { return "ELW" }
        30 { return "HMC" }
        36 { return "TWW" }
        default { return "UNK" }
    }
}

# Loads the System data. If nessecary grabs a fresh copy from the web
function LoadMainSystemDatabaseObject()
{
    $dbLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\StarSystemDataBase.xml"

    # Check to see if the db does not exists
    if(!(Test-Path $dbLocation))
    {
        # If the folder itself does not exist, create it
        if(!(Test-Path ($dbLocation.Replace("\StarSystemDataBase.xml", ""))))
        {
            New-Item -ItemType Directory -Path ($dbLocation.Replace("\StarSystemDataBase.xml", "")) | Out-Null
        }

        $inputFile = (Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&id=0B-_p8JFokEecZExFS3dxN0RzQ1U").Content | ConvertFrom-Json
        $readableFile = @()
        foreach($item in $inputFile.psobject.Properties)
        {
            $tempSys = NewSystemObject
 
            $tempSys.Name = $item.Name
            $tempSys.X = $item.Value.x
            $tempSys.Y = $item.Value.y
            $tempSys.Z = $item.Value.z
            $tempSys.ID = $item.Value.sid
 
            foreach($planet in $item.Value.planets.psobject.Properties)
            {
                $tempPlanet = NewPlanetObject
 
                $tempPlanet.Name = $planet.Name
                $tempPlanet.Distance = $planet.Value.dls
                $tempPlanet.Type = TranslateTypeCode $planet.Value.tp
                $tempPlanet.ID = $planet.Value.bid
 
                $tempSys.Planets += $tempPlanet
            }
 
            $readableFile += $tempSys
        }
 
        $readableFile | Export-Clixml $dbLocation
    }

    # assuming no error happened the file Exists, so load it
    return Import-Clixml $dbLocation
}

# Loads the user settings, or creates one if needed
function LoadMainUserSettingsObject()
{
    $userSettingsLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\UserSettings.xml"

    # Check to see if the db does not exists
    if(!(Test-Path $userSettingsLocation))
    {
        # If the folder itself does not exist, create it
        if(!(Test-Path ($userSettingsLocation.replace("\UserSettings.xml",""))))
        {
            New-Item -ItemType Directory -Path ($userSettingsLocation.replace("\UserSettings.xml", "")) | Out-Null
        }
        NewUserSettingsObject | Export-Clixml $userSettingsLocation
    }

    # assuming no error happened the file Exists, so load it
    return Import-Clixml $userSettingsLocation
}

# Saves the System data.
function SaveMainSystemDatabaseObject()
{
    $dbLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\StarSystemDataBase.xml"

    $mainSystemDatabaseObject | Export-Clixml $dbLocation

}

# Saves the UserSettingsObject
function SaveMainUserSettingsObject()
{
    $userSettingsLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\UserSettings.xml"

    $mainUserSettingsObject | Export-Clixml $userSettingsLocation
}
 
#endregion /* XML Communicator Functions */

#region /* Helper Functions */

# gets the number of systems left in a route
function GetSystemsLeft()
{
    $count = 0

    foreach($item in $mainUserSettingsObject.CurrentRoute)
    { 
        if($mainUserSettingsObject.SkipVisited)
        {
            if(!$item.Visited)
            {
                $count++
            } 
        }
        else
        {
            $count++
        }
    }

    return "$($count) - Left"
}

# finds a system in the main DB
function FindSystemInDataBase($sysName)
{
    $returnItem = $null

    foreach($sys in $mainSystemDatabaseObject)
    {
        if($sys.Name -eq $sysName)
        {
            $returnItem = $sys
            break
        }
    }

    return $returnItem
}

# sets the display to the current item in the route
function UpdateDisplayToCurrentData()
{
    $mainUserSettingsObject.CurrentSystem = $mainUserSettingsObject.CurrentRoute[0]
    $numberLeftLabel.Text = "$(GetSystemsLeft)" 
    $systemLabel.Text = $mainUserSettingsObject.CurrentSystem.Name
    $systemVisitedCheckbox.Checked = $mainUserSettingsObject.CurrentSystem.Visited
    $planetListDataGridView.DataSource = $mainUserSettingsObject.CurrentSystem.Planets   
}

# finds and replaces the given system in the db
function FindAndReplaceSystem($systemObject)
{
    foreach($sys in $mainSystemDatabaseObject)
    {
        if($sys.Name -eq $systemObject.Name)
        {
            $sys.Visited = $systemObject.Visited
            foreach($sOPlanet in $systemObject.Planets)
            {
                foreach($sPnt in $sys.Planets)
                {
                    if($sOPlanet.Name -eq $sPnt.Name)
                    {
                        $sPnt.Visited = $sOPlanet.Visited
                    }
                }
            }
        }
    }
}


#endregion /* Helper Functions */

#region /* Button Actions */ 

# runs the import button action
function RunImportButtonAction()
{
    $mainUserSettingsObject.CurrentRoute = @()

    $tempInput = $importTextBox.Text -split '[\r\n]'
    foreach($item in $tempInput)
    {
        if($item -match '\d+\s+\d+\.\d+\s')
        {
            $mainUserSettingsObject.CurrentRoute += FindSystemInDataBase "$(($item -split '\d+\s+\d+\.\d+\s')[1])"
        }
    }

    UpdateDisplayToCurrentData

    if($mainUserSettingsObject.CurrentSystem.Name -ne "No System Loaded")
    {
        $mainUserSettingsObject.CurrentSystem.Name | clip.exe
    }
}

# runs the Next button action
function RunNextButtonAction()
{
    if($mainUserSettingsObject.CurrentRoute.Count -eq 0)
    {
        $numberLeftLabel.Text = "No Route Entered"
        return
    }
    $mainUserSettingsObject.AutomarkSystem = $automarkStarCheckBox.Checked
    $mainUserSettingsObject.AutomarkPlanets = $automarkPlanetsCheckBox.Checked
    $mainUserSettingsObject.SkipVisited = $skipVisitedCheckBox.Checked

    #update the system in the DB
    if($mainUserSettingsObject.AutomarkPlanets)
    {
        foreach($item in $mainUserSettingsObject.CurrentSystem.Planets)
        {
            $item.Visited = $true
        }
    }
    else
    {
        foreach($dataPlanet in $planetListDataGridView.Rows)
        {
            foreach($item in $mainUserSettingsObject.CurrentSystem.Planets)
            {
                if($item.Name -eq $dataPlanet.Cells[0].Value)
                {
                    $item.Visited = $dataPlanet.Cells[4].Value
                }
            }
        }
    }

    if($mainUserSettingsObject.AutomarkSystem)
    {
        $mainUserSettingsObject.CurrentSystem.Visited = $true
    }
    else
    {
        $mainUserSettingsObject.CurrentSystem.Visited = $systemVisitedCheckbox.Checked
    }


    FindAndReplaceSystem $mainUserSettingsObject.CurrentSystem

    $mainUserSettingsObject.CurrentRoute = $mainUserSettingsObject.CurrentRoute | Where-Object { $_.Name -ne $mainUserSettingsObject.CurrentRoute[0].Name } 
    
    if($mainUserSettingsObject.SkipVisited)
    {
        while($mainUserSettingsObject.CurrentRoute.Count -gt 0)
        {
            if(!$mainUserSettingsObject.CurrentRoute[0].Visited)
            {
                $mainUserSettingsObject.CurrentSystem = $mainUserSettingsObject.CurrentRoute[0]
                break
            }
            else
            {
               $mainUserSettingsObject.CurrentRoute = $mainUserSettingsObject.CurrentRoute | Where-Object { $_.Name -ne $mainUserSettingsObject.CurrentRoute[0].Name } 
            }
        }
    }

    if($mainUserSettingsObject.CurrentRoute.Count -eq 0)
    {
        $mainUserSettingsObject.CurrentSystem.Name = "No System Loaded"
        $mainUserSettingsObject.CurrentSystem.X = ""
        $mainUserSettingsObject.CurrentSystem.Y = ""
        $mainUserSettingsObject.CurrentSystem.Z = ""
        $mainUserSettingsObject.CurrentSystem.ID = ""
        $mainUserSettingsObject.CurrentSystem.Visited = $false
        $mainUserSettingsObject.CurrentSystem.Planets = @()
    }

    UpdateDisplayToCurrentData

    $mainUserSettingsObject.CurrentSystem.Name | Clip.exe
}

# runs the Exit button Action
function RunExitButtonAction()
{
    $mainUserSettingsObject.AutomarkSystem = $automarkStarCheckBox.Checked
    $mainUserSettingsObject.AutomarkPlanets = $automarkPlanetsCheckBox.Checked
    $mainUserSettingsObject.SkipVisited = $skipVisitedCheckBox.Checked
    $mainUserSettingsObject.CurrentSystem.Visited = $systemVisitedCheckbox.Checked

    SaveMainSystemDatabaseObject
    SaveMainUserSettingsObject

}

# runs when the skipVisitedCheckBox is clicked
function OnSkipVisitedCheckBoxClick()
{
    $mainUserSettingsObject.SkipVisited = $skipVisitedCheckBox.Checked
    $numberLeftLabel.Text = "$(GetSystemsLeft)" 
}

# runs when a cell in the planets grid is clicked
function RunPlanetCellClicked()
{
    if($_.ColumnIndex -eq 4)
    {
        $planetListDataGridView.Rows[$_.RowIndex].Cells[4].Value = !($planetListDataGridView.Rows[$_.RowIndex].Cells[4].Value)
    }
}

#endregion /* Button Actions */
 
#region /* Windows Forms Functions */
 
function MainGUIConstructor()
{
    $mainForm.FormBorderStyle = 'Fixed3D'
    $mainForm.MaximizeBox = $false
    $mainForm.KeyPreview = $true
    $mainForm.ClientSize = New-Object System.Drawing.Size(640,177)
    $mainForm.Text = "ED - Road to Riches Helper"
    $mainForm.Icon = $iconPath
    $mainForm.Add_FormClosing({RunExitButtonAction})

    $systemLabel.Text = $mainUserSettingsObject.CurrentSystem.Name
    $systemLabel.Size = New-Object System.Drawing.Size(300,30)
    $systemLabel.Location = New-Object System.Drawing.Point(5,5)
    $systemLabel.Font = $labelFontBig
    $mainForm.Controls.Add($systemLabel)

    $systemVisitedCheckbox.Location = New-Object System.Drawing.Point(365,7)
    $systemVisitedCheckbox.Size = New-Object System.Drawing.Size(15,25)
    $systemVisitedCheckbox.Checked = $mainUserSettingsObject.CurrentSystem.Visited
    $mainForm.Controls.Add($systemVisitedCheckbox)
    $sysVisitedLabel = New-Object System.Windows.Forms.Label
    $sysVisitedLabel.Text = "- Visited"
    $sysVisitedLabel.Size = New-Object System.Drawing.Size(55,25)
    $sysVisitedLabel.Location = New-Object System.Drawing.Point(380,10)
    $sysVisitedLabel.Font = $labelFont
    $mainForm.Controls.Add($sysVisitedLabel)

    $planetListDataGridView.Size = New-Object System.Drawing.Size(430,137)
    $planetListDataGridView.Location = New-Object System.Drawing.Point(5,35)
    $planetListDataGridView.DataSource = $mainUserSettingsObject.CurrentSystem.Planets
    $planetListDataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $planetListDataGridView.Add_CellClick({RunPlanetCellClicked})
    $planetListDataGridView.ReadOnly = $true
    $mainForm.Controls.Add($planetListDataGridView)

    $mainDividerGroupBox = New-Object System.Windows.Forms.GroupBox
    $mainDividerGroupBox.Size = New-Object System.Drawing.Size(2,167)
    $mainDividerGroupBox.Location = New-Object System.Drawing.Point(445,5)
    $mainForm.Controls.Add($mainDividerGroupBox)

    $importTextBox.Size = New-Object System.Drawing.Size(100,25)
    $importTextBox.Location = New-Object System.Drawing.Point(455,7)
    $importTextBox.Multiline = $true
    $importTextBox.Text = "$($scriptComboBox.SelectedItem.Description)"
    $mainForm.Controls.Add($importTextBox)

    $importButton.Size = New-Object System.Drawing.Size(75,25)
    $importButton.Location = New-Object System.Drawing.Point(560,7)
    $importButton.Text = "Import"
    $importButton.Add_Click({RunImportButtonAction})
    $importButton.Font = $labelFont
    $mainForm.Controls.Add($importButton)

    $automarkStarCheckBox.Location = New-Object System.Drawing.Point(455,37)
    $automarkStarCheckBox.Size = New-Object System.Drawing.Size(15,25)
    $automarkStarCheckBox.Checked = $mainUserSettingsObject.AutomarkSystem
    $mainForm.Controls.Add($automarkStarCheckBox)
    $automarkStarLabel = New-Object System.Windows.Forms.Label
    $automarkStarLabel.Text = "- Automark System Visited"
    $automarkStarLabel.Size = New-Object System.Drawing.Size(200,25)
    $automarkStarLabel.Location = New-Object System.Drawing.Point(470,40)
    $automarkStarLabel.Font = $labelFont
    $mainForm.Controls.Add($automarkStarLabel)

    $automarkPlanetsCheckBox.Location = New-Object System.Drawing.Point(455,62)
    $automarkPlanetsCheckBox.Size = New-Object System.Drawing.Size(15,25)
    $automarkPlanetsCheckBox.Checked = $mainUserSettingsObject.AutomarkPlanets
    $mainForm.Controls.Add($automarkPlanetsCheckBox)
    $automarkPlanetsLabel = New-Object System.Windows.Forms.Label
    $automarkPlanetsLabel.Text = "- Automark Planets Visited"
    $automarkPlanetsLabel.Size = New-Object System.Drawing.Size(200,25)
    $automarkPlanetsLabel.Location = New-Object System.Drawing.Point(470,65)
    $automarkPlanetsLabel.Font = $labelFont
    $mainForm.Controls.Add($automarkPlanetsLabel)

    $skipVisitedCheckBox.Location = New-Object System.Drawing.Point(455,87)
    $skipVisitedCheckBox.Size = New-Object System.Drawing.Size(15,25)
    $skipVisitedCheckBox.Checked = $mainUserSettingsObject.SkipVisited
    $skipVisitedCheckBox.Add_Click({OnSkipVisitedCheckBoxClick})
    $mainForm.Controls.Add($skipVisitedCheckBox)
    $skipVisitedLabel = New-Object System.Windows.Forms.Label
    $skipVisitedLabel.Text = "- Skip Visited Systems"
    $skipVisitedLabel.Size = New-Object System.Drawing.Size(200,25)
    $skipVisitedLabel.Location = New-Object System.Drawing.Point(470,90)
    $skipVisitedLabel.Font = $labelFont
    $mainForm.Controls.Add($skipVisitedLabel)

    $numberLeftLabel.Text = "$(GetSystemsLeft)"
    $numberLeftLabel.Size = New-Object System.Drawing.Size(200,25)
    $numberLeftLabel.Location = New-Object System.Drawing.Point(455,117)
    $numberLeftLabel.Font = $labelFontBig
    $mainForm.Controls.Add($numberLeftLabel)

    $nextButton.Size = New-Object System.Drawing.Size(75,25)
    $nextButton.Location = New-Object System.Drawing.Point(455,147)
    $nextButton.Text = "Next"
    $nextButton.Add_Click({RunNextButtonAction})
    $nextButton.Font = $labelFont
    $mainForm.Controls.Add($nextButton)

    $exitButton.Size = New-Object System.Drawing.Size(75,25)
    $exitButton.Location = New-Object System.Drawing.Point(560,147)
    $exitButton.Text = "Exit"
    $exitButton.Add_Click({$mainForm.Close()})
    $exitButton.Font = $labelFont
    $mainForm.Controls.Add($exitButton)

}
 
#endregion /* Windows Forms Functions */
 
<#--- Main Script Function ---#>

 HideConsole

$mainSystemDatabaseObject = LoadMainSystemDatabaseObject

$mainUserSettingsObject = LoadMainUserSettingsObject

# copy the first Item to the clip board if it is not empty
if($mainUserSettingsObject.CurrentSystem.Name -ne "No System Loaded")
{
    $mainUserSettingsObject.CurrentSystem.Name | clip.exe
}
 
MainGUIConstructor
 
$mainForm.ShowDialog()
