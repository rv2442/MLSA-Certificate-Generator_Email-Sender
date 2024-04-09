# For troubleshooting only

# Get the running Outlook processes
$outlookProcesses = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue

# Check if there are any running Outlook processes
if ($outlookProcesses.Count -gt 0) {
    # Terminate all running Outlook processes
    $outlookProcesses | ForEach-Object {
        $_.CloseMainWindow()  # Attempt to close the main window gracefully
        if (-not $_.HasExited) {
            $_.Kill()  # Forcefully terminate the process if it hasn't exited gracefully
        }
    }
    Write-Host "All running Outlook processes terminated."
} else {
    Write-Host "No running Outlook processes found."
}
