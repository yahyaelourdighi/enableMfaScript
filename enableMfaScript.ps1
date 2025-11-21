Import-Module ImportExcel -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

$excelFile = "users.xlxs" # excel file containing users!
$logFile = "mfa_log.txt"

# Initialize log file with timestamp
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"=== MFA Configuration Log - $timestamp ===" | Out-File -FilePath $logFile -Encoding UTF8
"" | Out-File -FilePath $logFile -Append

# Connect to Microsoft Graph with required permissions
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All" -ErrorAction Stop
    Write-Host "✅ Successfully connected to Microsoft Graph" -ForegroundColor Green
    "✅ Successfully connected to Microsoft Graph" | Out-File -FilePath $logFile -Append
    "" | Out-File -FilePath $logFile -Append
}
catch {
    $errorMsg = "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Host $errorMsg -ForegroundColor Red
    $errorMsg | Out-File -FilePath $logFile -Append
    exit
}

# Import Excel file
$data = Import-Excel -Path $excelFile

# Use the specific column name "mail*"
$emailColumn = "mailColumn"

# Verify the column exists
if (-not ($data[0].PSObject.Properties.Name -contains $emailColumn)) {
    $errorMsg = "❌ Could not find the column '$emailColumn' in the Excel file!"
    Write-Host $errorMsg
    $errorMsg | Out-File -FilePath $logFile -Append
    Disconnect-MgGraph
    exit
}

Write-Host "Found email column: $emailColumn"
"Found email column: $emailColumn" | Out-File -FilePath $logFile -Append
"" | Out-File -FilePath $logFile -Append

# Counters for summary
$successCount = 0
$failureCount = 0

# Loop through each email
foreach ($user in $data) {
    $userId = $user.$emailColumn
    
    if ([string]::IsNullOrWhiteSpace($userId)) { 
        continue 
    }
    
    $userId = $userId.Trim()  # Trim spaces
    
    Write-Host "Enabling MFA for user: $userId"
    
    $body = @{ "perUserMfaState" = "enabled" }
    
    try {
        Invoke-MgGraphRequest -Method PATCH -Uri "/beta/users/$userId/authentication/requirements" -Body $body
        
        $successMsg = "✅ SUCCESS: MFA enabled for $userId"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
        $successCount++
    }
    catch {
        $failureMsg = "❌ FAILED: $userId - $($_.Exception.Message)"
        Write-Host $failureMsg -ForegroundColor Red
        $failureMsg | Out-File -FilePath $logFile -Append
        $failureCount++
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Write summary to log
"" | Out-File -FilePath $logFile -Append
"=== Summary ===" | Out-File -FilePath $logFile -Append
"Total successful: $successCount" | Out-File -FilePath $logFile -Append
"Total failed: $failureCount" | Out-File -FilePath $logFile -Append
"Total processed: $($successCount + $failureCount)" | Out-File -FilePath $logFile -Append

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host "Total successful: $successCount" -ForegroundColor Green
Write-Host "Total failed: $failureCount" -ForegroundColor Red
Write-Host "Total processed: $($successCount + $failureCount)"
Write-Host ""
Write-Host "Log file saved to: $logFile" -ForegroundColor Yellow