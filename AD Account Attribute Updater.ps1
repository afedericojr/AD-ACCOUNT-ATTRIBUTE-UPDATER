Write-Host "#################################################################" -f yellow
Write-Host "#                                                               #`n`r#" -nonewline -f yellow
Write-Host "                 AD ACCOUNT ATTRIBUTE UPDATER                  " -nonewline -f white
Write-Host "#`n`r#                                                               #
#  It is important to follow the directions closely, as         #
#  changes are permanent.                                       #
#                                                               #
#  This script will use a CSV file containing your user's       #
#  email addresses and desired attribute values to update       #
#  the specified profile attribute.                             #
#                                                               #
#  The CSV file must contain two columns.                       #
#  Set column 1 to Email                                        #
#  Set column 2 to AttributeValue                               #
#                                                               #
#  The log files will be saved in the directory of the script.  #
#                                                               #
#################################################################" -f yellow
Write-Host "
Copyright 2016 - Andrew J. Federico, Jr. - afedericojr@gmail.com

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE."
#>
# Initialize variables.
$title
$message
$yMessage
$nMessage
$actionType
$boolCheckYesNo = $false
$global:tryAgain = $false
$error.clear()

# Reusable function for confirming the user's desired input.
function checkYesNo([string]$actionType,[string]$title,[string]$message,[string]$yMessage,[string]$nMessage,[bool]$boolValue){
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Confirm and continue."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "Do not continue."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    if($result -eq 0){
        Write-Host $yMessage -f green
        $global:tryAgain = $boolValue
    }
    else{
        Write-Host $nMessage -f green
        if($actionType -eq "exit"){exit}
        else{}
    }
}

# First Y/N message stating to proceed with caution.
if(-not $boolCheckYesNo){checkYesNo "exit" "WARNING" "`n`rDo you agree to the use and licensing terms?" "`n`rProceed with caution" "`n`rGoodbye" $false}
Write-Host "`n`rENTER THE EXACT PATH TO THE OU IN AD." -f red
Write-Host "An example of this is: OU=New York,OU=Offices,DC=MyCompanyDomain,DC=com" -f yellow

# Set and confirm Active Directory connection string
Do{
    $treePath = Read-Host "`n`rWhat is your path?"
    checkYesNo "break" "IMPORTANT" "`n`rDid you enter it correctly?" "`n`rYour connection string has been set." "`n`rPlease enter the correct path." $true
}
Until ($global:tryAgain -eq $true)

# Reset the tryAgain boolean
$global:tryAgain = $false

# Set the logfiles from the name of the last OU in the connection string.
$logFileName = ($treePath -split ',')[1].substring(3) + " " + ($treePath -split ',')[0].substring(3)
$logFilePathSuccess = "$($logFileName) success.log"
$logFilePathError = "$($logFileName) error.log"

# Capture the CSV file, and retry on error.
Do{
    $userFile = Read-Host "`n`rWhat is the path to your CSV? If it is in the same directory of the script, simply enter the filename only."
    Try{
        $users = Import-CSV $userFile
        $global:tryAgain = $true
    }
    Catch{
        Write-Host = $_.Exception.Message -f red
        Write-Host = "`n`rPlease check the name of the file and try again." -f green
    }
}
Until ($global:tryAgain -eq $true)

# Reset the tryAgain boolean
$global:tryAgain = $false

# Set and confirm name of Active Directory profile attribute.
Do{
    $profileAttribute = Read-Host "`n`rWhat is the name of the attribute you wish to update?"
    checkYesNo "break" "IMPORTANT" "`n`rDid you enter it correctly?" "`n`rYour attribute selection has been made." "`n`rPlease enter the correct attribute name." $true
}
Until ($global:tryAgain -eq $true)

# Reset the tryAgain boolean
$global:tryAgain = $false

# Import each $user object (email address) from the CSV file, and use each email to select the user's SAM account, and update the attribute specified.
Do{
    $programStatus = "started"
    Try{
        ForEach ($user in $users) {
        $userEmail = $user.Email
        $userName = Get-ADUser -Filter {
            (EmailAddress -eq $userEmail)
            } -SearchBase $treePath -Properties EmailAddress |
        Select-Object sAMAccountName,EmailAddress
            if ($userName){
                Set-ADUser –Identity $userName.sAMAccountName -Clear "$profileAttribute" #COMMENT OUT FOR DEMO MODE
                if($user.AttributeValue){
                    Set-ADUser -Identity $userName.sAMAccountName -Add @{$profileAttribute = $user.AttributeValue} #COMMENT OUT FOR DEMO MODE
                    Write-Host "`n`rThe attribute for $($userEmail) has been set to $($user.AttributeValue)."  -f green
                    Write-Output "$($userEmail) has been successfully updated." | Out-File -FilePath $logFilePathSuccess -append -encoding ASCII
                    $programStatus = "completed"
                }
                else{
                    Write-Host "`n`rThe attribute for $($userEmail) is now cleared."  -f green
                    Write-Output "$($userEmail) has been successfully cleared." | Out-File -FilePath $logFilePathSuccess -append -encoding ASCII
                    $programStatus = "completed"
                    }
                    
            }
            else{
                Write-Host "`n`rThe script did not complete for user: $($userEmail)" -f red
                Write-Output "The script did not complete for user: $($userEmail)" | Out-File -FilePath $logFilePathError -append -encoding ASCII
                $programStatus = "completed"
            }
        }
    }
    Catch [System.Management.Automation.CommandNotFoundException]{
        Write-Host "`n`rThe command $($_.TargetObject) was not found!" -f red
            if($_.TargetObject -eq "Get-ADUser" -and !$global:tryAgain){
                # Set and confirm name of Active Directory profile attribute.
                    checkYesNo "exit" "IMPORTANT" "`n`rWould you like to install $($_.TargetObject) and try again." "`n`rInstalling $($_.TargetObject)...please wait." "`n`rYou have chosen not to install $($_.TargetObject)." $true
            }
            else{
                Write-Host "`n`r$($_.TargetObject) did not install correctly.`n`rThe script will now exit.`n`r" -f green
                $programStatus = "completed"
            }
    }
    Catch [System.ArgumentException]{
        Write-Host "`n`rThere was an error with parameters entered.`n`rPlease check the column names of your CSV file.`n`rEnsure that the columns are Email and AttributeValue respectively." -f green
        $programStatus = "completed"
    }
    Catch [Microsoft.ActiveDirectory.Management.ADException]{
        # Set and confirm name of Active Directory profile attribute.
        Do{
            $profileAttribute = Read-Host "`n`rPlease enter a valid attribute you wish to update?"
            checkYesNo "break" "IMPORTANT" "`n`rDid you enter it correctly?" "`n`rYour attribute selection has been made." "`n`rPlease enter the correct attribute name." $true
        }
        Until ($global:tryAgain -eq $true)

        # Reset the tryAgain boolean
        $global:tryAgain = $false
    }
    Catch {
        Write-Host $_.Exception.Message -f red
        Write-Host "`n`rThere was an unhandled error.`n`rPlease check your configuration and try again.`n`rFor your convenience a detailed error is shown below." -f green
        echo $_.Exception|format-list -force
        $programStatus = "completed"
    }
}
Until($programStatus -eq "completed")