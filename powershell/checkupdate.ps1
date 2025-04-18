# PowerShell Script for Windows Updates

try {
    # Create Update Session and Searcher objects
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()

    # Search for available software updates
    $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")

    # Check if updates are found
    if ($searchResult.Updates.Count -eq 0) {
        Write-Verbose "No updates to install." -Verbose # Using verbose for better output control
    }
    else {
        Write-Verbose "$($searchResult.Updates.Count) updates to install." -Verbose

        # Iterate through updates and display information
        foreach ($update in $searchResult.Updates) {
            Write-Verbose "Title: $($update.Title)" -Verbose
            Write-Verbose "Description: $($update.Description)" -Verbose
            Write-Verbose "IsDownloaded: $($update.IsDownloaded)" -Verbose
            Write-Verbose "------------------" -Verbose
        }
    }
}
catch {
    # Handle errors gracefully
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    # Cleanup (if needed) - COM objects are usually handled by garbage collection, but you could explicitly release them if you want.
    # [System.Runtime.InteropServices.Marshal]::ReleaseComObject($updateSearcher)
    # [System.Runtime.InteropServices.Marshal]::ReleaseComObject($updateSession)
}