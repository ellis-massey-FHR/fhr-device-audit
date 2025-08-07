# Define paths
$scriptPath = "C:\Users\ellism3\fhr-device-audit\scripts\get_servicenow_data.py"
$pythonExe = "python"  # Or use full path if needed
$workingDir = "C:\Users\ellism3\fhr-device-audit\scripts"

# Activate virtual environment (if needed)
& "$workingDir\venv\Scripts\Activate.ps1"

# Run the Python script
Write-Host "Running ServiceNow report script..."
cd $workingDir
& $pythonExe $scriptPath

# Open the Excel file
$excelFile = "$workingDir\filtered_laptop_requests.xlsx"
if (Test-Path $excelFile) {
    Write-Host "Opening report..."
    Start-Process $excelFile
} else {
    Write-Host "❌ Report not found!"
}

