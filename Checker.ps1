Set-StrictMode -Version Latest
Install-Package -name 'MimeKit' -Source "https://www.nuget.org/api/v2" -skipDependencies
Install-Package -name 'MailKit' -Source "https://www.nuget.org/api/v2" 


try {
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\System.Runtime.CompilerServices.Unsafe.6.1.2\lib\netstandard2.0\System.Runtime.CompilerServices.Unsafe.dll"
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\System.Memory.4.6.3\lib\netstandard2.0\System.Memory.dll"
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\System.Buffers.4.6.1\lib\netstandard2.0\System.Buffers.dll"

    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\BouncyCastle.Cryptography.2.6.2\lib\netstandard2.0\BouncyCastle.Cryptography.dll"
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.14.0\lib\netstandard2.0\MimeKit.dll"
    Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.14.1\lib\netstandard2.0\MailKit.dll"
} catch {
     Write-Error "Failed to load assemblies. Specific error: $($_.Exception.LoaderExceptions)"
}


$ComputerNamePath = "C:\Users\junio\Desktop\Powershell Projects\Enviorment Checker\EnviormentChecker.csv"
$ComputerNameList = Import-Csv -Path $ComputerNamePath


#Get-Credential | Export-Clixml -Path 'C:\Users\junio\Desktop\Powershell Projects\Enviorment Checker\gmail.xml'
$Account = Import-Clixml -Path "C:\Users\junio\Desktop\Powershell Projects\Enviorment Checker\gmail.xml"

Foreach($Computer in $ComputerNameList) {
    if(Test-Connection -TargetName $Computer.ComputerName -Count 1 -Quiet ) {
        Write-Host "Success: $Computer responded to ping."
    } else {
        Write-Host "Error: $($Computer.ComputerName) could not be contacted"

# ------------------------------------ Create 3 objects from Mimekit and MailKit  ------------------------------------------------------
        #1. Creating the smtp client to send email / Creating a "postal worker" object that knows how to talk to an email server.
        $SMTP    = New-Object MailKit.Net.Smtp.SmtpClient 
        #2. is the envelope and the letter inside it.
        $Message = New-Object MimeKit.MimeMessage
        #3. is where you can put your message in regards to the topic for the receiver
        $Builder = New-Object MimeKit.BodyBuilder

        # 2. Setup message details by accessing the objects attributes
        $Message.From.Add("juniorvalerioperdomo@gmail.com")
        $Message.To.Add("Jvp.Java@gmail.com")
        $Message.Subject = "CRITICAL: Environment Check - $($Computer.ComputerName) is Offline"

        # 3. Setup body details by accessing the objects attributes
        $MessageAlert = @"   
$([char]0x2022) Device Name: $($Computer.ComputerName) `n$([char]0x2022) Status: OFFLINE / UNREACHABLE `n$([char]0x2022) Time of Check: $(Get-Date -Format "h:mm tt") `n    
Description: The system was unable to contact the computer listed above during the scheduled environment check. This may indicate a network interruption, 
a power failure, or a system crash. `n 
                
Please investigate the connectivity status of this machine immediately. `n     
"@
        $Builder.TextBody = $MessageAlert
        $Message.Body = $Builder.ToMessageBody()

                # 3. Connect and send with corrected settings
        try {
            # Fix: Use 'smtp.gmail.com' and Port 587 
            # Connecting to gmail's smtp server with our client
            $SMTP.Connect('smtp.gmail.com', 587, $false) 
            
            # Authenticate with your 16-character App Password
            $SMTP.Authenticate($Account)
            $SMTP.Send($Message)

            Write-Host "Email sent successfully!"
        }
        catch {
            Write-Error "Failed to send: $_"
        }
        finally {
            # 4. Clean up carefully
            if ($SMTP.IsConnected) { $SMTP.Disconnect($true) }
            $SMTP.Dispose()
        }
    }
}




