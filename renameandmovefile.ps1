Install-Module -Name PSExcel

# Import the module
Import-Module PSExcel

# Set the source directory and pattern to match CSV files
$sourceDirectory = '/Users/max.hazell/downloads'
$filePattern = '*.csv'

# Get a list of CSV files in the source directory
$listOfFiles = Get-ChildItem -Path $sourceDirectory -Filter $filePattern

# Find the newest file based on creation time
$newest = $listOfFiles | Sort-Object CreationTime -Descending | Select-Object -First 1

# Display the path of the newest file
Write-Host $newest.FullName

$destinationDirectory = '/Users/max.hazell/Coding/hibob_ad_sync'
$destinationFilename = 'input.csv'
$destinationPath = Join-Path -Path $destinationDirectory -ChildPath $destinationFilename

# Copy the newest CSV file to the destination
Copy-Item -Path $newest.FullName -Destination $destinationPath

Import-Csv -Path '/Users/max.hazell/Coding/hibob_ad_sync/input.csv' | 
    Select-Object -Property @(
        @{ Name = 'Name';     Expression = 'Display Name' }
        @{ Name = 'Job Title';     Expression = 'Job Title' }
        @{ Name = 'Team';     Expression = 'Team/IPT' }
        @{ Name = 'PerManagerName';     Expression = 'Reports to' }
        @{ Name = 'WorkEmailAddress';     Expression = 'Email' }
    
    ) | 
    Export-Csv -Path '/Users/max.hazell/Coding/hibob_ad_sync/main.csv' -NoTypeInformation -UseQuotes AsNeeded
