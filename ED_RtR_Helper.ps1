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
    Add-Type -TypeDefinition @"
    
    public class PlanetObject
    {
        public string Name;
        public string Distance;
        public string Type;
        public string ID;
        public bool Visited;
    }
"@
  <#  
    $emptyPlanetObject= [PSCustomObject]@{
        Name = ""
        Distance = ""
        Type = ""
        ID = ""
        Visited = $false
    }
    #>
 
    return New-Object PlanetObject
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

    # DEBUG

        $emptyUserSettingsObject.CurrentSystem.Planets += NewPlanetObject
        $emptyUserSettingsObject.CurrentSystem.Planets += NewPlanetObject

        $emptyUserSettingsObject.CurrentSystem.Planets[0].Name = "Test 1"
        $emptyUserSettingsObject.CurrentSystem.Planets[1].Name = "Test 2"

        $emptyUserSettingsObject.CurrentSystem.Planets[1].Visited = $true

    # END DEBUG

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
$mainSystemDatabaseObject

# The user settings object
$mainUserSettingsObject
 
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
        if(!(Test-Path ($dbLocation -replace "\StarSystemDataBase.xml", "")))
        {
            New-Item -ItemType Directory -Path ($dbLocation -replace "\StarSystemDataBase.xml", "") | Out-Null
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
    $script:mainSystemDatabaseObject = Import-Clixml $dbLocation
}

# Loads the user settings, or creates one if needed
function LoadMainUserSettingsObject()
{
    $userSettingsLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\UserSettings.xml"

    # Check to see if the db does not exists
    if(!(Test-Path $userSettingsLocation))
    {
        # If the folder itself does not exist, create it
        if(!(Test-Path ($userSettingsLocation -replace "\\UserSettings.xml", "")))
        {
            New-Item -ItemType Directory -Path ($userSettingsLocation -replace "\\UserSettings.xml", "") | Out-Null
        }

        NewUserSettingsObject | Export-Clixml $userSettingsLocation
    }

    # assuming no error happened the file Exists, so load it
    $script:mainUserSettingsObject = Import-Clixml $userSettingsLocation
}

# Saves the System data.
function SaveMainSystemDatabaseObject()
{
    $dbLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\StarSystemDataBase.xml"

    $script:mainSystemDatabaseObject | Export-Clixml $dbLocation

}

# Saves the UserSettingsObject
function SaveUserSettingsObject()
{
    $userSettingsLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\UserSettings.xml"

    $script:mainUserSettingsObject | Export-Clixml $userSettingsLocation
}
 
#endregion /* XML Communicator Functions */

 
#region /* Windows Forms Functions */
 
function MainGUIConstructor()
{
    $mainForm.FormBorderStyle = 'Fixed3D'
    $mainForm.MaximizeBox = $false
    $mainForm.KeyPreview = $true
    $mainForm.Add_KeyDown({MainFormKeyDown})
    $mainForm.ClientSize = New-Object System.Drawing.Size(850,500)
    $mainForm.Text = "ED - Road to Riches Helper"
    $mainForm.Icon = $iconPath
    $mainForm.Add_FormClosing({SaveMainSystemDatabaseObject; LoadMainUserSettingsObject})

    $systemLabel.Text = $mainUserSettingsObject.CurrentSystem.Name
    $systemLabel.Size = New-Object System.Drawing.Size(300,30)
    $systemLabel.Location = New-Object System.Drawing.Point(5,5)
    $systemLabel.Font = $labelFontBig
    $mainForm.Controls.Add($systemLabel)

    $systemVisitedCheckbox.Location = New-Object System.Drawing.Point(310,7)
    $systemVisitedCheckbox.Size = New-Object System.Drawing.Size(20,25)
    $systemVisitedCheckbox.Checked = $mainUserSettingsObject.CurrentSystem.Visited
    $mainForm.Controls.Add($systemVisitedCheckbox)
    $sysVisitedLabel = New-Object System.Windows.Forms.Label
    $sysVisitedLabel.Text = "- Visited"
    $sysVisitedLabel.Size = New-Object System.Drawing.Size(100,25)
    $sysVisitedLabel.Location = New-Object System.Drawing.Point(330,10)
    $sysVisitedLabel.Font = $labelFont
    $mainForm.Controls.Add($sysVisitedLabel)

    $planetListDataGridView.Size = New-Object System.Drawing.Size(430,250)
    $planetListDataGridView.Location = New-Object System.Drawing.Point(5,35)
    $planetListDataGridView.DataSource = $mainUserSettingsObject.CurrentSystem.Planets
    $mainForm.Controls.Add($planetListDataGridView)

}
 
#endregion /* Windows Forms Functions */
 
<#--- Main Script Function ---#>

#HideConsole

LoadMainSystemDatabaseObject

LoadMainUserSettingsObject
 
MainGUIConstructor
 
$mainForm.ShowDialog()
