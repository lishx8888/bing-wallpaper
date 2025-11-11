# Bing Wallpaper Fix Script
# A reliable script to download and set Bing daily wallpaper

Write-Host "Starting Bing Wallpaper Fix Script..."

# Configuration
$saveFolder = "D:\TD\Pictures\BingWallpapers"
$today = Get-Date -Format "yyyyMMdd"
$imageName = "Bing_$today.jpg"
$imagePath = Join-Path -Path $saveFolder -ChildPath $imageName
$logFile = "$PSScriptRoot\bing_wallpaper_fix_log.txt"
$waitForUpdate = $true

# Log function - defined before being used
function Add-LogEntry($message) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $message"
    Write-Host $logLine
    try {
        $logLine | Add-Content -Path $logFile -ErrorAction SilentlyContinue
    } catch {
        # Silently continue if logging fails
    }
}

# Clean up wallpapers older than 30 days
function CleanupOldWallpapers {
    try {
        $cutoffDate = (Get-Date).AddDays(-30)
        $filesToDelete = Get-ChildItem -Path $saveFolder -Filter "Bing_*.jpg" -File | 
                        Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($filesToDelete.Count -gt 0) {
            Write-Host "Cleaning up $($filesToDelete.Count) wallpapers older than 30 days..."
            Add-LogEntry -message "Cleaning up $($filesToDelete.Count) wallpapers older than 30 days"
            
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force
                Add-LogEntry -message "Deleted old wallpaper: $($file.FullName)"
            }
            
            Write-Host "Cleanup completed successfully."
        } else {
            Write-Host "No wallpapers older than 30 days found."
        }
    } catch {
        Write-Host "Error during cleanup: $($_.Exception.Message)"
        Add-LogEntry -message "Error during cleanup: $($_.Exception.Message)"
    }
}

# Create save folder if it doesn't exist
if (-not (Test-Path -Path $saveFolder)) {
    try {
        New-Item -Path $saveFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created wallpaper save directory: $saveFolder"
        Add-LogEntry -message "Created wallpaper save directory: $saveFolder"
    } catch {
        Write-Host "Error creating directory: $($_.Exception.Message)"
        Add-LogEntry -message "Error creating directory: $($_.Exception.Message)"
    }
}

# Run cleanup before downloading new wallpaper
CleanupOldWallpapers

# Wait for 8 AM update if needed
if ($waitForUpdate) {
    $currentTime = Get-Date
    $targetTime = Get-Date -Hour 8 -Minute 0 -Second 0
    
    if ($currentTime -lt $targetTime) {
        $waitMinutes = [math]::Round(($targetTime - $currentTime).TotalMinutes, 0)
        Write-Host "Current time is before 8 AM, waiting $waitMinutes minutes for latest Bing wallpaper..."
        Add-LogEntry -message "Waiting for Bing wallpaper update, current time: $currentTime, target time: $targetTime"
        # Uncomment the next line if you want to actually wait
        # Start-Sleep -Seconds ($targetTime - $currentTime).TotalSeconds
    }
}

# Check if today's wallpaper already exists
if (Test-Path -Path $imagePath) {
    Write-Host "Today's wallpaper already exists, setting it directly..."
    Add-LogEntry -message "Today's wallpaper already exists, setting directly: $imageName"
} else {
    # Download Bing wallpaper
    Write-Host "Downloading today's Bing wallpaper..."
    Add-LogEntry -message "Starting to download today's Bing wallpaper"
    
    try {
        $bingUrl = "https://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
        Write-Host "Fetching image info from: $bingUrl"
        
        # Use -UseBasicParsing to avoid IE dependencies
        $response = Invoke-RestMethod -Uri $bingUrl -Method Get -UseBasicParsing -ErrorAction Stop
        
        if ($response.images -and $response.images.Count -gt 0) {
            $imageInfo = $response.images[0]
            $imageUrl = "https://cn.bing.com" + $imageInfo.urlbase + "_1920x1080.jpg"
            $copyright = $imageInfo.copyright
            
            Write-Host "Downloading image from: $imageUrl"
            Write-Host "Image description: $copyright"
            
            # Download the image
            Invoke-WebRequest -Uri $imageUrl -OutFile $imagePath -UseBasicParsing -ErrorAction Stop
            
            Write-Host "Image downloaded successfully to: $imagePath"
            Add-LogEntry -message "Successfully downloaded wallpaper: $imageName - $copyright"
        } else {
            Write-Host "Error: No images found in Bing API response!"
            Add-LogEntry -message "Error: No images found in Bing API response"
            Exit 1
        }
    } catch {
        Write-Host "Error: Failed to download Bing wallpaper!"
        Write-Host "Error details: $($_.Exception.Message)"
        Add-LogEntry -message "Error: Failed to download Bing wallpaper: $($_.Exception.Message)"
        Exit 1
    }
}

# Set the wallpaper using Windows API (most reliable method)
Write-Host "Setting wallpaper..."
Add-LogEntry -message "Setting wallpaper: $imagePath"

try {
    # Add Windows API method for setting wallpaper
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    
    public class WallpaperChanger
    {
        public const uint SPI_SETDESKWALLPAPER = 0x0014;
        public const uint SPIF_UPDATEINIFILE = 0x01;
        public const uint SPIF_SENDCHANGE = 0x02;
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(
            uint uAction,
            uint uParam,
            string lpvParam,
            uint fuWinIni);
            
        public static bool SetWallpaper(string wallpaperPath)
        {
            int result = SystemParametersInfo(
                SPI_SETDESKWALLPAPER,
                0,
                wallpaperPath,
                SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
            
            return result != 0;
        }
    }
"@
    
    # Set wallpaper style via registry first
    Write-Host "Updating registry settings for wallpaper style..."
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value 1  # Fill
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value 0   # Not tiled
    
    # Ensure the image file exists
    if (-not (Test-Path -Path $imagePath)) {
        throw "Image file not found: $imagePath"
    }
    
    # Convert to absolute path (important for API)
    $absoluteImagePath = (Resolve-Path -Path $imagePath).Path
    Write-Host "Using absolute path: $absoluteImagePath"
    
    # Use Windows API to set wallpaper
    Write-Host "Calling Windows API to set wallpaper..."
    $success = [WallpaperChanger]::SetWallpaper($absoluteImagePath)
    
    if (-not $success) {
        throw "Windows API failed to set wallpaper"
    }
    
    # Additional refresh to ensure changes apply
    Write-Host "Performing final desktop refresh..."
    Start-Process -FilePath "rundll32.exe" -ArgumentList "user32.dll,UpdatePerUserSystemParameters" -Wait
    
    Write-Host "Wallpaper successfully set! Please check your desktop."
    Add-LogEntry -message "Wallpaper successfully set: $imagePath"
} catch {
    Write-Host "Error: Failed to set wallpaper!"
    Write-Host "Error details: $($_.Exception.Message)"
    Add-LogEntry -message "Error: Failed to set wallpaper: $($_.Exception.Message)"
    Exit 1
}

Write-Host "Bing Wallpaper Fix Script completed successfully!"
Add-LogEntry -message "Script completed successfully"