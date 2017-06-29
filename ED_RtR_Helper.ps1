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
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
 
[DllImport("user32.dll")]
public static extern void DisableProcessWindowsGhosting();
 
'
 
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
    $emptySystemObject= New-Object psobject -Property @{
        "Name" = ""
        "X" = ""
        "Y" = ""
        "Z" = ""
        "ID" = ""
        "Visited" = $false
        "Planets" = @()
    }
 
    return $emptySystemObject
}
 
# declares a new planet object.
function NewPlanetObject()
{
    $emptyPlanetObject= New-Object psobject -Property @{
        "Name" = ""
        "Distance" = ""
        "Type" = ""
        "ID" = ""
        "Visited" = $false
    }
 
    return $emptyPlanetObject
}
 
#endregion  /* Script Data Objects */
 
#region /* Global Form Objects */
 
# the windows forms objects
$mainForm = New-Object System.Windows.Forms.Form

# the default icon path
$iconPath = ".\Data\DefaultIcon.ico"

# The main Data Objects for the script
$mainSystemDatabaseObject
 
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
    $mainSystemDatabaseObject = Import-Clixml $dbLocation
}

# Saves the System data.
function SaveMainSystemDatabaseObject()
{
    $dbLocation = "C:\Users\$($env:USERNAME)\AppData\Roaming\ED_RtR_Helper\StarSystemDataBase.xml"

    $mainSystemDatabaseObject | Export-Clixml $dbLocation

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
    $mainForm.Add_FormClosing({SaveSystemList})
}
 
#endregion /* Windows Forms Functions */
 
<#--- Main Script Function ---#>

#HideConsole

LoadMainSystemDatabaseObject
 
MainGUIConstructor
 
$mainForm.ShowDialog()
