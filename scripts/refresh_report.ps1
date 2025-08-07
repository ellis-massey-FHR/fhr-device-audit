# Define paths
$scriptPath = "C:\Users\ellism3\fhr-device-audit\scripts\get_servicenow_data.py"
$pythonExe = "C:\Users\ellism3\fhr-device-audit\venv\Scripts\python.exe"
$workingDir = "C:\Users\ellism3\fhr-device-audit"

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
