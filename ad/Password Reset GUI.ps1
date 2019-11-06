<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Password Reset Tool
.SYNOPSIS
    Provides a graphical user interface for non-IT staff to reset passwords and unlock accounts
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '450,272'
$Form.text                       = "Password Reset"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Search"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(42,38)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$txtSearch                       = New-Object system.Windows.Forms.TextBox
$txtSearch.multiline             = $false
$txtSearch.width                 = 100
$txtSearch.height                = 20
$txtSearch.location              = New-Object System.Drawing.Point(102,34)
$txtSearch.Font                  = 'Microsoft Sans Serif,10'

$lblFullName                     = New-Object system.Windows.Forms.Label
$lblFullName.text                = "Full Name"
$lblFullName.AutoSize            = $true
$lblFullName.width               = 25
$lblFullName.height              = 10
$lblFullName.location            = New-Object System.Drawing.Point(229,38)
$lblFullName.Font                = 'Microsoft Sans Serif,10'

$btnSearch                       = New-Object system.Windows.Forms.Button
$btnSearch.text                  = "Search"
$btnSearch.width                 = 60
$btnSearch.height                = 30
$btnSearch.location              = New-Object System.Drawing.Point(44,81)
$btnSearch.Font                  = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Password"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(44,162)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$txtPassword                     = New-Object system.Windows.Forms.TextBox
$txtPassword.multiline           = $false
$txtPassword.width               = 100
$txtPassword.height              = 20
$txtPassword.enabled             = $false
$txtPassword.location            = New-Object System.Drawing.Point(129,163)
$txtPassword.Font                = 'Microsoft Sans Serif,10'

$btnReset                        = New-Object system.Windows.Forms.Button
$btnReset.text                   = "Reset"
$btnReset.width                  = 60
$btnReset.height                 = 30
$btnReset.enabled                = $false
$btnReset.location               = New-Object System.Drawing.Point(45,206)
$btnReset.Font                   = 'Microsoft Sans Serif,10'

$btnClear                        = New-Object system.Windows.Forms.Button
$btnClear.text                   = "Clear"
$btnClear.width                  = 60
$btnClear.height                 = 30
$btnClear.enabled                = $false
$btnClear.location               = New-Object System.Drawing.Point(321,206)
$btnClear.Font                   = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Label1,$txtSearch,$lblFullName,$btnSearch,$Label3,$txtPassword,$btnReset,$btnClear))


#region main
$global:cred = Get-Credential
$global:user = $null
function resetControls()
{
    $txtSearch.text = $null
    $lblFullName.text = "FirstName LastName"
    $txtPassword.text = $null
    $btnReset.enabled = $false
    $btnClear.enabled = $false
    $global:user = $null
}

function changePassword($length)
{
    $password = generateComplexPassword($length)
    $title = "Random Password"
    $message = "The value $password has been copied to your clipboard. Please communicate the change as this password is not stored anywhere."
    $buttons = "Ok"
    [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
    Set-Clipboard -Value $password
    $txtPassword.text = $password
    
    $from = 'IT Support <itsupport@contoso.com'
    $to = '$global:user.Name $global:user.UserPrincipalName'
    $cc = '$env:username' + '@' + '$env:USERDNSDOMAIN'
    $subject = 'Password Reset Request has come in'
    $body = 'Your password has been reset to' + ' ' + $txtPassword.text
    $smtpserver = 'smtp.office365.com'
    $smtpport = '587'
    
    $mailparams = @{
        From = $from
        To = $to
        cc = $cc
        Subject = $subject
        Body = $body
        SMTPServer = $smtpserver
        Port = $smtpport
    }
    
    #Start-Sleep 2
    #Send-MailMessage @mailparams -UseSSL -Credential $global:cred
}

$btnSearch.add_Click(
 {
    If ([string]::IsNullOrWhiteSpace($txtSearch.Text))
    {
        $message = "Please enter a valid username"
        $title = "Informational Warning"
        $buttons = "Ok"
        [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
        $txtSearch.text = $null
    }
    ElseIf (userExists($txtSearch.text))
    {
        $global:user = Get-ADUser -Identity $txtSearch.text -Properties *
        $lblFullName.text = $global:user.CN
        Switch ($global.user)
        {
            {$global:user.Enabled -ne $true}# scriptblocks evaluate expressions
            {
                $message = "User's account is disabled. Please contact support"
                $title = "Informational Warning"
                $buttons = "Ok"
                [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                break
            }
            {$global:user.LockedOut -eq $true}
            {
                $message = "User's account is locked. Do you wish to unlock"
                $title = "Action Warning"
                $buttons = "YesNo"
                $flag = [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                If ($flag -eq "Yes")
                {
                    $message = "Your account has been unlocked!"
                    $title = "Informational Window"
                    $buttons = "Ok"
                    [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                }
                Else
                {
                    $message = "Your account remains locked!"
                    $title = "Informational Window"
                    $buttons = "Ok"
                    [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                }
            }
            {$global:user.PasswordExpired -eq $true}
            {
                $message = "User's password has expired. Please reset it"
                $title = "Action Warning"
                $buttons = "YesNo"
                $flag = [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                If ($flag -eq "Yes")
                {
                    # Set-ADAccountPassword -Identity $global:user -Reset -NewPassword changePassword(8) -Credential $global:cred
                    $message = "Your password has been updated!"
                    $title = "Informational Window"
                    $buttons = "Ok"
                    [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                }
                Else
                {
                    $message = "Your password remains the same!"
                    $title = "Informational Window"
                    $buttons = "Ok"
                    [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
                }
            }
            default
            {
                $btnReset.enabled = $true
                $btnClear.enabled = $true
                $message = "You made it through!"
                $title = "Informational Warning"
                $buttons = "Ok"
                [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
            }
        }
    }
    Else
    {
        $title = "Informational Warning"
        $message = "User does not exist in AD"
        $buttons = "Ok"
        [System.Windows.Forms.MessageBox]::Show($message,$title,$buttons)
        $txtSearch.text = "Username"
        $lblFullName.text = "FirstName LastName"
        resetControls
    }
 })

$btnReset.add_Click(
{
   changePassword(8)
})

$btnClear.add_Click(
{
    resetControls
})
#endregion
#region boolean
function userExists ($user)
{
    $flag = [bool] (Get-ADUser -Filter { SamAccountName -eq $user })
    If ($flag -eq $true)
    {
        return $true
    }
    Else
    {
        return $false
    }
    
}

function isLockedOut ($user)
{
    cls
    Try
    {
        $flag = Get-ADUser -Identity $user -Properties * | Select LockedOut
        return $flag.LockedOut
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }

}

function isEnabled ($user)
{
    cls
    Try
    {
        $flag = Get-ADUser -Identity $user -Properties * | Select Enabled
        return $flag.Enabled
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }

}

function isPasswordExpired ($user)
{
    cls
    Try
    {
        $flag = Get-ADUser -Identity $user -Properties * | Select PasswordExpired
        return $flag.PasswordExpired
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }

}
#endregion

#region actions

function Enable($user)
{
    cls
    Try
    {
        $account = Get-ADUser -Identity $user -Properties * 
        If ($account.Enabled -eq $true)
        {
            Write-Host "User is already enabled!" -ForegroundColor Yellow
        }
        else
        {
            $mycred = $host.ui.PromptForCredential("Authorization Prompt", `
            "Please enter your admin account and password to complete this action",$admin,"")
            if ($mycred.Password -eq $null)
            {
                Write-Host $noValue -ForegroundColor Yellow
            }
            else
            {
                Enable-ADAccount -Identity $user -Credential $mycred
                Write-Host "$($account.Givenname)'s account enabled!" -ForegroundColor Green
            }   
        }
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }
    Catch [System.Security.Authentication.AuthenticationException]
    {
        Write-Host $notValidorNotAllowed -ForegroundColor Red
    }

}

function Disable($user)
{
    cls
    Try
    {
        $account = Get-ADUser -Identity $user -Properties *
        If ($account.Enabled -eq $false)
        {
            Write-Host "User is already disabled!" -ForegroundColor Yellow
        }
        else
        {
            $mycred = $host.ui.PromptForCredential("Authorization Prompt", `
            "Please enter your admin account and password to complete this action",$admin, "")
            if ($mycred.Password -eq $null)
            {
                Write-Host $noValue -ForegroundColor Yellow
            }
            else
            {
                Disable-ADAccount -Identity $user -Credential $mycred
                Write-Host "$($account.Givenname)'s account Disabled!" -ForegroundColor Green
            }
            
        } 
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }
    Catch [System.Security.Authentication.AuthenticationException]
    {
        Write-Host $notValidorNotAllowed -ForegroundColor Red
    }

}

function ResetPassword ($user) {

    cls
    Try
    {
        $mycred = $host.ui.PromptForCredential("Authorization Prompt", `
        "Please enter your admin account and password to complete this action",$admin, "")
        $account = Get-ADUser -Identity $user -Properties *
        $newpass = Read-Host -AsSecureString -Prompt "Please enter a strong password"
        if ($mycred.Password -eq $null)
        {
            Write-Host $noValue -ForegroundColor Yellow
        }
        else
        {
            Set-ADAccountPassword -Identity $user -Reset -NewPassword $newpass -Credential $mycred
            Write-Host "$($account.Givenname)'s password has been changed!" -ForegroundColor Green
        }
        
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }
    Catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException]
    {
        Write-Host "ERROR! password does not meet complexity requirements!" -ForegroundColor Red
    }
    Catch [Microsoft.ActiveDirectory.Management.ADInvalidPasswordException]
    {
        Write-Host "ERROR! password is invalid" -ForegroundColor Red
    }
    Catch [System.Security.Authentication.AuthenticationException]
    {
        Write-Host $notValidorNotAllowed -ForegroundColor Red
    }
}

function Unlock ($user)
{
    cls
    Try
    {
        $account = Get-ADUser -Identity $user -Properties *
        If ($account.LockedOut -eq $false)
        {
            $name = $account.GivenName
            Write-Host "$name's account is already unlocked!" -ForegroundColor Yellow
        }
        else
        {
            $mycred = $host.ui.PromptForCredential("Admin Prompt", `
            "Please enter your admin account and password to complete this action",$admin,"")
            if ($mycred.Password -eq $null)
            {
                Write-Host $noValue -ForegroundColor Yellow
            }
            else
            {
                Unlock-ADAccount -Identity $user -Credential $mycred
                Write-Host "$($account.Givenname)'s account unlocked!" -ForegroundColor Green
            }
        }
    }
    Catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException]
    {
        Write-Host $notfoundinAD -ForegroundColor Red
    }
    Catch [System.Security.Authentication.AuthenticationException]
    {
        Write-Host $notValidorNotAllowed -ForegroundColor Red
    }
}

function generateComplexPassword ($length)
{
    $digits = 48..57
    $uppercase = 65..90
    $lowercase = 97..122
    $special = 33..42
    $password = $null
    $atLeastOneUpperCase = "[A-Z]+"
    $atLeastOneLowerCase = "[a-z]+"
    $atLeastOneNumber = "[0-9]+"
    $match = $atLeastOneUpperCase + $atLeastOneLowerCase + $atLeastOneNumber -join ""

    $combined = ([char[]]$digits) + ([char[]]$uppercase) + ([char[]]$lowercase) + ([char[]]$special)

    do 
    {
        $password = (Get-Random -Count $length -InputObject([char[]]$combined)) -join ""
    }
    until ($password -match $match) #Loops until string has at least one uppercase, lowercase and number

    return $password
}
#endregion
Clear-Host
[void]$Form.ShowDialog()