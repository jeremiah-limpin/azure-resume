# Define variables
$speedtestexe = "C:\speedtest\speedtest.exe"
$computerName = $env:COMPUTERNAME
$resultsFilePath = "C:\$computerName_speedtest_results.txt"

# Download Speedtest CLI
if (Test-Path $speedtestexe) {
    Write-Host "File is present, moving on to testing."
} else {
    Invoke-WebRequest -Uri "https://install.speedtest.net/app/cli/ookla-speedtest-1.0.0-win64.zip" -OutFile "C:\speedtest.zip"
    Expand-Archive -Path "C:\speedtest.zip" -DestinationPath "C:\speedtest"
}

# Run Speedtest and output results
& $speedtestexe --accept-license --accept-gdpr > $resultsFilePath

# Get the public IP address
$ip = Invoke-RestMethod -Uri "http://ifconfig.me/ip"

# Create an Outlook email to send the results via Microsoft Exchange
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Namespace.Logon()  # Log on to the default profile (automatically uses default account)

    # Create the mail item
    $Mail = $Outlook.CreateItem(0)  # 0 means a new mail item

    # Check if the mail item was created successfully
    if ($Mail -eq $null) {
        Write-Host "Failed to create mail item."
        return
    }

    # Configure the email
    $Mail.Subject = "Speedtest CLI Results for $computerName"
    $Mail.Body = "Please find the attached Speedtest CLI results from $computerName."
    $Mail.To = "helpdesk@thebackroomop.com"  # Replace with the recipient's email address

    # Attach the Speedtest results file
    $Mail.Attachments.Add($resultsFilePath)

    # Send the email automatically
    $Mail.Send()
    Write-Host "Email with Speedtest results has been sent successfully."

} catch {
    # If there is an error, catch it and output an error message
    Write-Host "Failed to send email. Error: $_"
}