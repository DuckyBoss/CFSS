#THIS VERSION NOT TO BE USED OUTSIDE OF CFSS USE THE DEMO FILE

#----------------------Known Issues----------------------
#Windows activation, activation not giving correct response to computers that have a hardware change 

#When ran as admin Checking for windows activation raises a "Get-Object" Error
#Fixed - Why was I trying to run this script as admin?

#Windows update function wont check when there is no wifi 
#Fix - Make a function that tries to connect to the cfss wifi before checking for updates

#add processor type: Search for i3, i5, i7 in processor info in system info

#Check Wifi should try to connect to the cfss network when disconnected but still found the adpater

#When manufacturer is Hewlett Packard, chage to HP


#---------------------------------------------------------



#----------------------Things-To-Add----------------------
#Final screen - True is green False is red
#Percentage Completed as the functions are running, maybe a loading bar of some sort
#Auto win 10 activation with codes stored in a diffrent csv or json file
#Home page where user can select purpose: Run scan, Add win 10 codes
#Activation type I.E. Activated with Windows 10 Home
#Organize The spec list to match computer sheet
#Run internet speed test and report it by wifi aspect

#---------------------------------------------------------



#--------------------------------------------------------------------------






try{
    Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Unrestricted -Force
}
catch{
    Write-Host "Unable to Set execution Policy" -ForegroundColor red
    Pause
}
finally{
    Clear-Host
    $Info = systeminfo  
}




Write-Host "Loading..."




function Manufacturer {
    try {
        $SysInfo = [string]($Info | findstr /C:"System Manufacturer")

        $SysInfo = $SysInfo.Replace("System Manufacturer:", "")

        $SysInfo = $SysInfo.Replace("       ", "")
        $SysInfo = $SysInfo.Replace("      ", "")
        $SysInfo = $SysInfo.Replace("     ", "")
        $SysInfo = $SysInfo.Replace("    ", "")

        return $SysInfo
    }
    catch {
        return "error"
    }
        


}

function Model {
    try {
        $SysInfo = [string]($Info | findstr /C:"System Model")

        $SysInfo = $SysInfo.Replace("System Model:", "")

        $SysInfo = $SysInfo.Replace("       ", "")
        $SysInfo = $SysInfo.Replace("      ", "")
        $SysInfo = $SysInfo.Replace("     ", "")
        $SysInfo = $SysInfo.Replace("    ", "")

        return $SysInfo
    }
    catch {
        return "error"
    }
        


}

function RAM {
    try {
        $SysInfo = $Info | findstr /C:"Total Physical Memory"
        $Sysinfo = $SysInfo.replace("Total Physical Memory:     ", "")
        $SysInfo = ([int]$SysInfo.replace(" MB", "")) / 1000
        $SysInfo = [math]::Round($SysInfo)


        return $SysInfo
    }
    catch {
        return "error"
    }
        
}

function CPU_Speed {
    try {
        $SysInfo = Get-WmiObject Win32_Processor | findstr /C:"MaxClockSpeed"
        $SysInfo = $SysInfo.replace(" ", "")
        $SysInfo = $SysInfo.replace("MaxClockSpeed:", "")
        $SysInfo = [int]$SysInfo / 1000

        return $SysInfo
    }
    catch {
        return "error"
    }
    
}

function Storage {
    try {
        $SysInfo = [string](Get-Volume -DriveLetter C | Select-Object -Property Size)

        
        #Removes "@{Size="
        $SysInfo = $SysInfo.replace("@{Size=", "")
            
        #Removes last curly bracket
        $SysInfo = $SysInfo.replace("}", "")

        #it would be 12 digits if it was in the hundreds of gigabytes
        if ($SysInfo.Length -eq 12) {
            $SysInfo = $SysInfo.Substring(0, 3)
        }

        #It would be 13 digits if it was in the terabytes
        elseif ($SysInfo.Length -eq 13) {
            $SysInfo = $SysInfo.Substring(0, 4)
        }

        #Storage not suitable
        else {
            $SysInfo = "Not between 100 - 10,000 gb"
        }
            
        return $SysInfo
    }
    catch {
        return "error"
    }
}


function WiFi {

    function connectWiFi{
        #Main CFSS wifi with P@55word
        try{
            netsh wlan connect ssid="CFSS" key="P@55word"
            WirelessStatus
        }
        catch{
            return $false
        }
            
    }

    function WirelessStatus{
        try {
            $Internet_Settings = netsh interface show interface | findstr /C:"Wi-Fi" /C:"Name"
        
            if ($Internet_Settings -like "*Wi-Fi*" -and $Internet_Settings -like "*Connected*") {
                return $true
            }
            elseif ($Internet_Settings -like "*Wireless*" -and $Internet_Settings -like "*Connected*") {
                return $true
            }
        
        
            if ($Internet_Settings -like "*Wi-Fi*" -and $Internet_Settings -like "*Disconnected*") {
                connectWiFi
            }
            elseif ($Internet_Settings -like "*Wireless*" -and $Internet_Settings -like "*Disconnected*") {
                connectWiFi
            }    
        
            else {
                return $false
            }
        }
        catch {
            return "error"
        }
    }

        
}

function CD_Drive {
    try {
        $drives = Get-WmiObject Win32_Volume -Filter "DriveType=5"
        if ($null -eq $drives) {
        
            return "No CD Drive Detected"
        
        } 
        
        $drives | ForEach-Object {
        (New-Object -ComObject Shell.Application).Namespace(17).ParseName($_.Name).InvokeVerb("Eject")
        }
        
            
        function AskUser {
            Clear-Host
            $AskHost = Read-Host "Did the CD drive Eject? [y/n]"

            if ($AskHost -like "y") {
                return $true
            }
            elseif ($AskHost -like "n") {
                return "Drive Detected, Not ejected"
            }
            else {
                Clear-Host
                Write-Host "I have to write an entire function to handle people like you"
                Start-Sleep 3
                Clear-Host; Write-Host "Just write [y,n]"; Start-Sleep 2
                Clear-Host ; Write-Host "Lets try that again" ; Start-Sleep 1
                AskUser
            }
        }
        AskUser
            
    }
    catch {
        return "error"
    }

    
}

function Sound {
    #I would like to have the Start-Sleep in the finally section under the PlaySound function wait for the duration of the song but not sure how and im on a deadline

    Clear-Host


    function AskUser {
        Clear-Host
        $AskUser = Read-Host "Did you hear it?[y/n] ['r' repeat]"

        if ($AskUser -like "y") {
            return $true
        }
        elseif ($AskUser -like "n") {
            return $false
        }
        elseif ($AskUser -like "r") {
            PlaySound
        }
        else {
            Clear-Host
            Write-Host "I have to write an entire function to handle people like you"
            Start-Sleep 3
            Clear-Host; Write-Host "Just write [y,n] or [r] if you want the sound repeated"; Start-Sleep 4
            Clear-Host ; Write-Host "Lets try that again" ; Start-Sleep 1
            AskUser
        }
    }

    function PlaySound {
        Clear-Host
        Write-Host "Playing Sound"
            

        try {
            (New-Object System.Media.SoundPlayer $(Get-Random $(Get-ChildItem -Path "$env:windir\Media\Ring*.wav").FullName)).Play()
        }
        #If it couldnt find file itll play simple tone
        catch {
            [console]::beep(740, 1500)
        }
        finally {
            Start-Sleep 4
            Clear-Host
            Write-Host "Sound Played"
            AskUser
        }
            
    }

    PlaySound

        



        
}


function WindowsUpdate {
    try {
        $criteria = "Type='software' and IsAssigned=1 and IsHidden=0 and IsInstalled=0"

        $searcher = (New-Object -COM Microsoft.Update.Session).CreateUpdateSearcher()
        $updates = $searcher.Search($criteria).Updates
        
        if ($updates.Count -ne 0) {
            control update
            return $false              #Windows is not up to date
        }
        
        else {
            return $true                #Windows is up to date
        }
    }
    catch {
        return "error"
    }


}

function Activation {

        
        
    function ActivateWindows {
            

        try{
            $TxtLength = (Get-Content Win10_Activation_Codes.txt).Length
        }
        catch{
            Write-Host "Could Not find Win10_Activation_Code.txt"
            return $false
        }


        $x = 0

        while ($x -le $TxtLength) {

            #Finds the string contents corresponding to the line number
            try {
                $key = Get-Content Win10_Activation_Codes.txt | Select-Object -Index $x
            }
            catch {
                return "Error Pulling key from Win10_Activation_Codes.txt"
            }
                
            $check.InstallProductKey($key)
            $check.RefreshLicenseStatus()

            $x = $x + 1
        }
    }

    $check = get-wmiObject -query 'select * from SoftwareLicensingService'

    #This will run if windows is already activated
    if ($null -ne $check.OA3xOriginalProductKey) {
        #Write-Host "Already Liscenced With Key: $($check.OA3xOriginalProductKey)" -ForegroundColor Green
        return $true
        break
    }
    else {
        function ConfirmUser {
            Clear-Host
            $AskUser = Read-Host "Windows is Not activated, Would you like to activate it?"
            if ("*y*" -like $AskUser) {
                ActivateWindows
            }
            elseif ("*n*" -like $AskUser) {
                return $false
            }
            else {
                ConfirmUser
            }
        }
        ConfirmUser
    }

    $x = $x + 1

}



function main {
    
    Write-Host "Loading Manufacturer Information ..."
    $Manufacturer = Manufacturer
    Clear-Host

    Write-Host "Loading Model Information ..."
    $Model = Model
    Clear-Host

    Write-Host "Loading Total RAM Information ..."
    $RAM = RAM
    Clear-Host

    Write-Host "Loading Storage Information ..."
    $Storage = Storage
    Clear-Host

    Write-Host "Loading CPU Clock Speed Information ..."
    $ClockSpeed = CPU_Speed
    Clear-Host


    Write-Host "Loading WiFi Adapter ..."
    $WiFi = CheckWiFi
    Clear-Host

    Write-Host "Testing CD Drive ..."
    $CdDrive = CheckCD
    Clear-Host

    Write-Host "Testing Sound ..."
    $Sound = Sound
    Clear-Host

    Write-Host "Locating Windows Updates ..."
    $Update = WindowsUpdate
    Clear-Host

    Write-Host "Loading Windows Activation ..."
    $Activation = Activation
    Clear-Host




    Write-Host "System Specs`n---------------------" -ForegroundColor blue
    Write-Host "Manufacturer:" $Manufacturer
    Write-Host "Model:" $Model
    Write-Host "RAM:" $RAM "gb"
    Write-Host "Storage:" $Storage "gb"
    Write-Host "CPU Clock Speed:" $ClockSpeed "GHz"

    Write-Host "`nAdditional`n---------------------" -ForegroundColor blue
    Write-Host "WiFi:" $WiFi
    Write-Host "CD Drive:" $CdDrive
    Write-Host "Sound:" $Sound
    Write-Host "Windows Activation:" $Activation
    Write-Host "Win10 Updated:" $Update

    Write-Host "`n"



    $TotalIssues = 0
    $IssueMessage = ""

    if (8 -gt $RAM) {
        $IssueMessage = $IssueMessage + "`nNot Enough RAM"
        $TotalIssues = $TotalIssues + 1
    }
    if ($WiFi -ne $true) {
        $IssueMessage = $IssueMessage + "`nNo Wifi Adapter Found" 
        $TotalIssues = $TotalIssues + 1
    }
    if ($CdDrive -ne $true) {
        $IssueMessage = $IssueMessage + "`n" + $CdDrive 
        $TotalIssues = $TotalIssues + 1
    }
    if ($Activation -ne $true) {
        $IssueMessage = $IssueMessage + "`nWindows Not Activated"
        $TotalIssues = $TotalIssues + 1
    }
    if ($Update -ne $true) {
        $IssueMessage = $IssueMessage + "`nWindows Is Not Updated" 
        $TotalIssues = $TotalIssues + 1
    }
    if ($Sound -ne $true) {
        $IssueMessage = $IssueMessage + "`nSound Not Working" 
        $TotalIssues = $TotalIssues + 1
    }


    if ($TotalIssues -eq 0) {
        Write-Host "Meets CFSS Standard!" -ForegroundColor green
    }
    else {
        Write-Host $TotalIssues "Issue(s) Found" -NoNewline
        Write-Host $IssueMessage -ForegroundColor red
    }

    pause

}

main

