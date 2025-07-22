Get-ChildItem -Path $PSScriptRoot -Directory | ForEach-Object {
    $subdirectory = $_.FullName
    Write-Host "Processing directory: $subdirectory"

    try {
        Write-Host "Running command 1..."
        cd $subdirectory
        dotnet new console
        dotnet add package DocumentFormat.OpenXml --version 2.20.0
        dotnet add package System.Drawing.Common --version 9.0.5
        dotnet add package SpreadsheetLight.Cross.Platform --version 3.7.0
        rm Program.cs
        dotnet build
        dotnet run
        cd ..
    }
    catch {
        Write-Error "An error occurred in directory $($_.FullName): $($_.Exception.Message)"
    }
}