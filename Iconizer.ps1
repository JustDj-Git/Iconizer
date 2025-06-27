function Shout {
    param(
        [parameter(Mandatory = $true)]
        [string]$text,
        [string]$color,
        [switch]$new,
        [switch]$after,
        [switch]$date
    )
    
    if (($date) -or ($log)){
        $_date = (Get-Date -Format "MM/dd/yy HH:mm:ss").ToString()
        $finaltext = "{0} {1}" -f $_date, $text
    } else {
        $finaltext = $text
    }
    
    if ($new){ $finaltext = "`n" + $finaltext }
    if ($after){ $finaltext = $finaltext + "`n" }
    if ($log) { $finaltext | Out-File -FilePath $log -Append -ErrorAction SilentlyContinue }
    
    if ($color){
        if (-not ([Enum]::IsDefined([System.ConsoleColor], $color))) {
            Write-Host "$color doesn't exist in System.ConsoleColor" -ForegroundColor Red
            Write-Host $finaltext
        } else {
            Write-Host $finaltext -ForegroundColor $color
        }
    } else {
        Write-Host $finaltext
    }
}

function Timer {
    param(
        [switch]$start,
        [switch]$end
    )
    if ($start){
        $global:timer = [Diagnostics.Stopwatch]::StartNew()
    }
    if ($end){
        $global:timer.Stop()
        $timeRound = [Math]::Round(($global:timer.Elapsed.TotalSeconds), 2)
        $global:timer.Reset()
        Shout "Task completed in $timeRound`s" -color Cyan -new
    }
}

function SelectPath {
    param(
        [switch]$files
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $Topmost = New-Object System.Windows.Forms.Form
    $Topmost.TopMost = $True
    $Topmost.MinimizeBox = $True
    
    if ($files){
        $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.RestoreDirectory = $True
        $OpenFileDialog.Title = 'Select an EXE File'
        $OpenFileDialog.Filter = 'Executable files (*.exe)|*.exe'
        if (($OpenFileDialog.ShowDialog($Topmost) -eq 'OK')) {
            $file = $OpenFileDialog.FileName
        } else {
            $file = $null
        }
    } else {
        $OpenFolderDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
        $OpenFolderDialog.Description = 'Select a folder'
        $OpenFolderDialog.Rootfolder = 'MyComputer'
        $OpenFolderDialog.ShowNewFolderButton = $false
        if ($OpenFolderDialog.ShowDialog($Topmost) -eq 'OK') {
            $directory = $OpenFolderDialog.SelectedPath
        } else {
            $directory = $null
        }
    }
    
    $Topmost.Close()
    $Topmost.Dispose()
    
    if ($file){
        return $file
    } elseif ($directory){
        return $directory
    } else {
        return $null
    }
}

function Import-Type {
    try {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class IconExtractor
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, uint dwFlags);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool FreeLibrary(IntPtr hModule);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr FindResourceW(IntPtr hModule, string lpName, IntPtr lpType);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr LockResource(IntPtr hResData);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);
    
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumResourceNames(IntPtr hModule, IntPtr lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);
    
    public delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);
    
    public const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
    public const int RT_GROUP_ICON = 14;
    public const int RT_ICON = 3;
    
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, uint nIconIndex);
    
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
    
    // Helper method to check if a pointer represents an integer resource
    public static bool IS_INTRESOURCE(IntPtr ptr)
    {
        return ((ulong)ptr) >> 16 == 0;
    }
    
    // Method to get the main icon index (simplified approach)
    public static int GetMainIconIndex(string filePath)
    {
        try
        {
            // Get total icon count
            uint iconCount = ExtractIconEx(filePath, -1, null, null, 0);
            return iconCount > 0 ? 0 : -1; // Return 0 for first icon, -1 if no icons
        }
        catch
        {
            return -1;
        }
    }
}
"@

    # Add .NET assemblies for image processing
    Add-Type -AssemblyName System.Drawing

    } catch [System.Exception] {
        Write-Host "An unexpected error occurred in Import-Type: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

function Get-IconsByGroup {
    param(
        [string]$FilePath,
        [int]$index = 1,
        [string]$OutputDir = ".",
        [switch]$all,
        [switch]$info,
        [switch]$png
    )
    
    if (-not (Test-Path $FilePath)) {
        Shout "File not found: $FilePath" -color Red
        return
    }
    
    if (-not $info -and -not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    
    if ($info) {
        $all = $true
    }
    
    $ICO_name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    
    $hModule = [IconExtractor]::LoadLibraryEx($FilePath, [IntPtr]::Zero, [IconExtractor]::LOAD_LIBRARY_AS_DATAFILE)
    
    if ($hModule -eq [IntPtr]::Zero) {
        Shout "Failed to load file!" -color Red
        return
    }
    
    try {
        $script:currentGroup = 0
        $script:targetGroup = $index
        $script:extractAll = $all
        $script:totalExtracted = 0
        $script:processedGroups = @()
        $script:resourcesNames = @()
        
        $message = if ($script:extractAll) {
            'Analyzing all icon groups'
        } else {
            "Analyzing group #$index"
        }
        
        Shout "$message..." -color Yellow -new
        
        $callback = {
            param($hMod, $lpType, $lpName, $lParam)
            
            $script:currentGroup++
            
            # Variables to track the best icon from current group
            $currentGroupLargestIcon = $null
            $currentGroupLargestSize = 0
            
            if (-not $script:extractAll -and $script:currentGroup -ne $script:targetGroup) {
                return $true
            }
            
            # Determine resource name/ID
            if ([IconExtractor]::IS_INTRESOURCE($lpName)) {
                $resourceId = [int]$lpName
                $resourceName = "ID_$resourceId"
            } else {
                try {
                    $stringName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($lpName)
                    $resourceName = if ([string]::IsNullOrWhiteSpace($stringName)) { "UNNAMED" } else { $stringName }
                } catch {
                    $resourceName = "ERROR_READING_NAME"
                }
            }
            
            Shout "Extracting group #$script:currentGroup ($resourceName)..." -color Green -new
            
            # Load and analyze icon group resource
            $hResInfo = [IntPtr]::Zero
            
            if ([IconExtractor]::IS_INTRESOURCE($lpName)) {
                $hResInfo = [IconExtractor]::FindResource($hMod, $lpName, [IntPtr][IconExtractor]::RT_GROUP_ICON)
            } else {
                try {
                    $stringName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($lpName)
                    if (-not [string]::IsNullOrEmpty($stringName)) {
                        $hResInfo = [IconExtractor]::FindResourceW($hMod, $stringName, [IntPtr][IconExtractor]::RT_GROUP_ICON)
                    }
                } catch {
                    Shout "Error processing string name" -color Red
                }
            }
            
            $groupExtracted = $false
            
            if ($hResInfo -ne [IntPtr]::Zero) {
                $hResData = [IconExtractor]::LoadResource($hMod, $hResInfo)
                if ($hResData -ne [IntPtr]::Zero) {
                    $pData = [IconExtractor]::LockResource($hResData)
                    $size = [IconExtractor]::SizeofResource($hMod, $hResInfo)
                    
                    if ($pData -ne [IntPtr]::Zero -and $size -gt 6) {
                        # Read group resource data
                        $iconDir = New-Object byte[] $size
                        [System.Runtime.InteropServices.Marshal]::Copy($pData, $iconDir, 0, $size)
                        
                        # Parse group icon header
                        $iconCount = [BitConverter]::ToUInt16($iconDir, 4)
                        
                        Shout "Icons found in group: $iconCount" -color Cyan
                        
                        # Create ICO file
                        if ($all){
                            $icoPath = Join-Path $OutputDir "${ICO_name}_Group_${script:currentGroup}_${resourceName}.ico"
                        } else {
                            $icoPath = Join-Path $OutputDir "$ICO_name.ico"
                        }
                        
                        $icoData = @()
                        
                        # ICO file header
                        $icoData += @(0, 0)  # Reserved
                        $icoData += @(1, 0)  # Type (1 = ICO)
                        $icoData += [BitConverter]::GetBytes([uint16]$iconCount)  # Count
                        
                        $iconDataArray = @()
                        $currentOffset = 6 + ($iconCount * 16)  # Header + directory
                        
                        # Process each icon in group
                        for ($i = 0; $i -lt $iconCount; $i++) {
                            $offset = 6 + ($i * 14)
                            if ($offset + 13 -lt $iconDir.Length) {
                                $width = $iconDir[$offset]
                                $height = $iconDir[$offset + 1]
                                $colorCount = $iconDir[$offset + 2]
                                $reserved2 = $iconDir[$offset + 3]
                                $planes = [BitConverter]::ToUInt16($iconDir, $offset + 4)
                                $bitCount = [BitConverter]::ToUInt16($iconDir, $offset + 6)
                                $iconId = [BitConverter]::ToUInt16($iconDir, $offset + 12)
                                
                                # Load individual icon data
                                $hIconRes = [IconExtractor]::FindResource($hMod, [IntPtr]$iconId, [IntPtr][IconExtractor]::RT_ICON)
                                if ($hIconRes -ne [IntPtr]::Zero) {
                                    $hIconData = [IconExtractor]::LoadResource($hMod, $hIconRes)
                                    if ($hIconData -ne [IntPtr]::Zero) {
                                        $pIconData = [IconExtractor]::LockResource($hIconData)
                                        $iconSize = [IconExtractor]::SizeofResource($hMod, $hIconRes)
                                        
                                        if ($pIconData -ne [IntPtr]::Zero -and $iconSize -gt 0) {
                                            $iconBytes = New-Object byte[] $iconSize
                                            [System.Runtime.InteropServices.Marshal]::Copy($pIconData, $iconBytes, 0, $iconSize)
                                            
                                            # Check if this is the largest icon in current group for PNG extraction
                                            if ($png) {
                                                $actualWidth = if ($width -eq 0) { 256 } else { $width }
                                                $actualHeight = if ($height -eq 0) { 256 } else { $height }
                                                $iconPixelSize = $actualWidth * $actualHeight
                                                
                                                if ($iconPixelSize -gt $currentGroupLargestSize) {
                                                    $currentGroupLargestSize = $iconPixelSize
                                                    $currentGroupLargestIcon = @{
                                                        Width = $actualWidth
                                                        Height = $actualHeight
                                                        Data = $iconBytes
                                                        Group = $script:currentGroup
                                                        ResourceName = $resourceName
                                                    }
                                                }
                                            }
                                            
                                            # Add icon directory to ICO file
                                            $icoData += @($width, $height, $colorCount, $reserved2)
                                            $icoData += [BitConverter]::GetBytes($planes)
                                            $icoData += [BitConverter]::GetBytes($bitCount)
                                            $icoData += [BitConverter]::GetBytes([uint32]$iconSize)
                                            $icoData += [BitConverter]::GetBytes([uint32]$currentOffset)
                                            
                                            $iconDataArray += ,$iconBytes
                                            $currentOffset += $iconSize
                                            if ($bitCount -eq 0) { $bitCount = 32 }
                                            if ($width -eq 0) { $width = 256 }
                                            if ($height -eq 0) { $height = 256 }
                                            Shout "  Icon $($i+1): ${width}x${height}, $bitCount bit, $iconSize bytes" -color Gray
                                        }
                                    }
                                }
                            }
                        }
                        
                        # Write ICO file
                        if ($iconDataArray.Count -gt 0) {
                            if (!($info) -and !($png)){
                                $allData = @()
                                $allData += $icoData
                                foreach ($iconBytes in $iconDataArray) {
                                    $allData += $iconBytes
                                }
                                [System.IO.File]::WriteAllBytes($icoPath, $allData)
                                Shout "Saved: $icoPath" -color Green
                            }
                            $script:totalExtracted++
                            $script:processedGroups += $script:currentGroup
                            $script:resourcesNames += $resourceName
                            $groupExtracted = $true
                        }
                    }
                }
            }
            
            if (-not $groupExtracted) {
                Shout "Failed to extract group #$script:currentGroup" -color Red
            }
            
            # Extract largest icon from current group as PNG if requested
            if ($png -and $currentGroupLargestIcon) {
                $pngFileName = if ($script:extractAll) {
                    "${ICO_name}_Group_${script:currentGroup}_${resourceName}.png"
                } else {
                    "${ICO_name}.png"
                }
                $pngPath = Join-Path $OutputDir $pngFileName
                
                Convert-IconToPNG -IconData $currentGroupLargestIcon -PngPath $pngPath -ICO_name $ICO_name
            }
            
            return $script:extractAll
        }
        
        $callbackDelegate = [IconExtractor+EnumResNameProc]$callback
        [IconExtractor]::EnumResourceNames($hModule, [IntPtr][IconExtractor]::RT_GROUP_ICON, $callbackDelegate, [IntPtr]::Zero) | Out-Null
        
        if ($script:totalExtracted -eq 0) {
            if ($script:extractAll) {
                Shout "No icon groups found or failed to extract any groups" -color Red
            } else {
                Shout "Group #$index not found or failed to extract" -color Red
            }
        } else {
            Shout "Extraction completed successfully!" -color Green -new -after
            Shout "Total groups extracted: $script:totalExtracted" -color Green
            Shout "Processed groups: $($script:resourcesNames -join ', ')" -color Cyan -after
        }
    }
    finally {
        [IconExtractor]::FreeLibrary($hModule) | Out-Null
    }
}

function Convert-IconToPNG {
    param(
        [hashtable]$IconData,
        [string]$PngPath,
        [string]$ICO_name
    )
    
    try {
        Shout "Extracting icon as PNG: $($IconData.Width)x$($IconData.Height) from Group $($IconData.Group)" -color Cyan
        
        $conversionSuccess = $false
        
        # Method 1: Try direct PNG extraction if the icon data is already PNG
        if ($IconData.Data.Length -gt 8) {
            $pngSignature = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
            $isPng = $true
            for ($i = 0; $i -lt 8; $i++) {
                if ($IconData.Data[$i] -ne $pngSignature[$i]) {
                    $isPng = $false
                    break
                }
            }
            
            if ($isPng) {
                Shout "Icon data is already PNG format, saving directly..." -color Cyan
                [System.IO.File]::WriteAllBytes($PngPath, $IconData.Data)
                $conversionSuccess = $true
            }
        }
        
        # Method 2: Try .NET conversion if not PNG
        if (-not $conversionSuccess) {
            try {
                # Create properly formatted ICO file
                $tempIcoData = @()
                $tempIcoData += @(0, 0, 1, 0, 1, 0)  # ICO header for single icon
                
                # Icon directory entry (16 bytes)
                $width = if ($IconData.Width -eq 256) { 0 } else { $IconData.Width }
                $height = if ($IconData.Height -eq 256) { 0 } else { $IconData.Height }
                $tempIcoData += @($width, $height, 0, 0)  # width, height, colors, reserved
                $tempIcoData += @(1, 0, 32, 0)  # planes, bitcount
                $tempIcoData += [BitConverter]::GetBytes([uint32]$IconData.Data.Length)  # size
                $tempIcoData += [BitConverter]::GetBytes([uint32]22)  # offset
                
                # Add icon data
                $tempIcoData += $IconData.Data
                
                # Save temporary ICO file
                $tempIcoPath = Join-Path $env:TEMP "temp_largest_icon.ico"
                [System.IO.File]::WriteAllBytes($tempIcoPath, $tempIcoData)
                
                Shout "Saving as png..." -color Cyan
                
                # Try multiple .NET approaches
                try {
                    # Approach 1: Direct Icon loading
                    Shout "Using direct Icon loading..." -color Cyan
                    $icon = [System.Drawing.Icon]::new($tempIcoPath)
                    $bitmap = $icon.ToBitmap()
                    $bitmap.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $bitmap.Dispose()
                    $icon.Dispose()
                    $conversionSuccess = $true
                } catch {
                    Shout "Direct Icon loading failed: $($_.Exception.Message)" -color Yellow
                    
                    # Approach 2: Try extracting from file stream
                    try {
                        Shout "Using FileStream approach..." -color Cyan
                        $fileStream = [System.IO.FileStream]::new($tempIcoPath, [System.IO.FileMode]::Open)
                        $icon = [System.Drawing.Icon]::new($fileStream)
                        $bitmap = $icon.ToBitmap()
                        $bitmap.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
                        $bitmap.Dispose()
                        $icon.Dispose()
                        $fileStream.Close()
                        $conversionSuccess = $true
                    } catch {
                        Shout "FileStream approach failed: $($_.Exception.Message)" -color Yellow
                    }
                }
                
                # Cleanup temp file
                Remove-Item $tempIcoPath -Force -ErrorAction SilentlyContinue
                
            } catch {
                Shout ".NET conversion failed: $($_.Exception.Message)" -color Yellow
            }
        }
        
        # Method 3: Save raw icon data as fallback
        if (-not $conversionSuccess) {
            Shout "Saving raw icon data as fallback..." -color Yellow
            $rawPath = $PngPath -replace '\.png$', '_raw.bin'
            [System.IO.File]::WriteAllBytes($rawPath, $IconData.Data)
            Shout "Raw icon data saved: $rawPath" -color Yellow
            Shout "You can try converting this file manually with image editing software" -color Yellow
        }
        
        if ($conversionSuccess) {
            Shout "Icon saved as PNG: $PngPath" -color Green
            return $true
        } else {
            Shout "PNG conversion failed, but raw data was saved for manual conversion" -color Yellow
            return $false
        }
        
    } catch {
        Shout "Error during PNG extraction: $($_.Exception.Message)" -color Red
        return $false
    }
}

function Find-Candidates {
    param (
        [parameter(Mandatory = $true)]
        [string]$path,
        [parameter(Mandatory = $true)]
        [int]$search_depth,
        [ValidateSet('ico', 'exe')]
        $priority
    )
    
    $folder = Get-Item -LiteralPath $path
    
    [string]$full_path_folder = $folder.FullName
    
    $iconsFilesList = Get-ChildItem -LiteralPath "$full_path_folder" -Recurse -Filter "*.ico" -Depth $search_depth -File
    $exesFilesList = Get-ChildItem -LiteralPath "$full_path_folder" -Recurse -Filter "*.exe" -Depth $search_depth -File
    $allFiles = @($exesFilesList) + @($iconsFilesList)
    
    [string]$name_folder = ($folder.Name).ToLower()
    $name_folder = $name_folder -replace $pattern_regex, '' -replace "$pattern_regex_digits", '' -replace "$pattern_regex_symbols", ''
    
    if ($allFiles) {
        if ($log) {
            Shout "Found: $allFiles" -color DarkGray
        }
        
        $candidates = @()
        
        $folderWords = $name_folder.Trim().Split(' ')
        $fullFolderName = $name_folder -replace '\s+', ''
        
        foreach ($file in $allFiles) {
            $FileName = ($file.BaseName).ToLower() -replace $pattern_regex_symbols, '' -replace $pattern_regex, '' -replace $pattern_regex_digits, ''
            $score = 0
            
            if ($FileName -eq $fullFolderName) {
                $score = 1000
            } elseif ($FileName.Contains($fullFolderName) -or $fullFolderName.Contains($FileName)) {
                $score = 500 + ($FileName.Length - [Math]::Abs($FileName.Length - $fullFolderName.Length))
            } else {
                foreach ($word in $folderWords) {
                    if ($word.Length -gt 2) {
                        if ($FileName -eq $word) {
                            $score += 200
                        } elseif ($FileName.Contains($word)) {
                            $score += 100
                        } elseif ($word.Contains($FileName) -and $FileName.Length -gt 3) {
                            $score += 50
                        }
                    }
                }
                
                if ($score -gt 0) {
                    $lengthDiff = [Math]::Abs($FileName.Length - $fullFolderName.Length)
                    $score += [Math]::Max(0, 20 - $lengthDiff)
                }
            }
            
            if ($priority -and $file.Extension.ToLower() -eq ".$priority") {
                $score += 400
                Shout "Priority bonus applied to: `'$($file.Name)`'" -color Cyan
            }
            
            if ($score -gt 0) {
                # Bonus for short names
                $score += [Math]::Max(0, 50 - $FileName.Length)
                $candidates += [PSCustomObject]@{
                    File  = $file
                    Score = $score
                    Name  = $FileName
                }
                if ($log){
                    Shout "Candidate: $($file.Name) | Score: $score" -color DarkGray
                }
            }
        }
        
        if ($candidates.Count -gt 0) {
            $bestCandidate = $candidates | Sort-Object Score -Descending | Select-Object -First 1
            $Files = $bestCandidate.File
            if ($log) {
                Shout "Best candidate: $($Files.Name) with score $($bestCandidate.Score)" -color DarkGray
            }
        } else {
            $Files = $allFiles | Select-Object -First 1
            Shout "No matches found, using first exe: `'$($Files.Name)`'" -color Yellow
        }
    } else {
        Shout ".$priority candidates not found" -color Red
    }
    return $Files
}

function Test-ForbiddenFolder {
    param (
        [string]$Path,
        [string[]]$ForbiddenFolders
    )
    
    $pathParts = $Path -split '\\' | Where-Object { $_ -ne '' }
    
    foreach ($forbiddenFolder in $ForbiddenFolders) {
        if ($pathParts -contains $forbiddenFolder) {
            return $true
        }
    }
    return $false
}

function pull {
    [CmdletBinding()]
    param (
        [Alias('d')]
        [string[]]$directory,
        [Alias('i')]
        [int]$index = 1,
        [Alias('dep')]
        [int]$depth = 0,
        [switch]$png,
        [switch]$info,
        [Alias('a')]
        [switch]$all,
        [switch]$folder,
        [switch]$pause,
        [Alias('l')]
        [string]$log
    )
    
    Timer -start
    
    Import-Type
    
    if (!($directory)) {
        if ($folder){
            $directory = SelectPath
        } else {
            $directory = SelectPath -files
        }
        
        if ($directory){
            $file_GUI = $true
        }
    }
    
    if (!($directory)) {
        Shout "No path was selected by the user. Select the path in the GUI or specify it like -d 'full_path_to_files'" -color Red -new -after
        return
    }
    
    $ErrorActionPreference = 'Stop'
    
    try {
        Shout "Let's start extracting icons from exes" -color Green -new
        Shout "-----------------"
        
        foreach ($i in $directory) {
            if (Test-Path $i){
                if ($file_GUI){
                    $resolved_path = Get-ChildItem -Path $i -Filter '*.exe'
                } else {
                    $resolved_path = Get-ChildItem -Path $i -Filter '*.exe' -Recurse -Depth $depth
                }
                
                foreach ($_path in $resolved_path){
                    if ($_path) {
                        Shout "Extracting icons from:`n $($_path.FullName)" -color Yellow -new
                        $params = @{
                            FilePath  = $_path.FullName
                            OutputDir = $_path.DirectoryName
                            index     = $index
                        }
                        
                        if ($all) { $params.all   = $true }
                        if ($info) { $params.info  = $true }
                        if ($png) { $params.png   = $true }
                        Get-IconsByGroup @params
                    } else {
                        Shout "No exe files in path:`n $($_path.FullName)" -color Red
                    }
                } #foreach
            } else {
                Shout "Path is not exist:`n $($_path.FullName)" -color Red
            }
        } #foreach
    } catch {
        Shout "Error:$_" -color Red -new
        Shout "$($_.ScriptStackTrace)" -color Red -new -after
    }
    
    Timer -end
    
    if ($pause){
        pause
    }
}

function apply {
    [CmdletBinding()]
    param (
        [Alias('d')]
        [string[]]$directory,
        [Alias('p')]
        [ValidateSet('ico', 'exe')]
        $priority,
        [Alias('f')]
        [string[]]$filter,
        [Alias('s')]
        [switch]$single,
        [Alias('rm')]
        [switch]$remove,
        [Alias('l')]
        [string]$log,
        [Alias('r')]
        $rules,
        [Alias('nf')]
        [switch]$NoForce,
        [switch]$pause,
        [Alias('sd')]
        [int]$search_depth = 0,
        [Alias('ad')]
        [int]$apply_depth = 0
    )
    
    try {
        if (!($directory)) {
            $directory = SelectPath
        }
        if (!($directory)) {
            Shout "No path was selected by the user. Select the path in the GUI or specify it like -d 'full_path_to_files'" -color Red -new -after
            return
        }
        
        $ErrorActionPreference = 'Stop'
        $foldersError = @()
        $folders = @()
        Shout "Let's start applying icons to folders" -color Green -new
        Shout "-----------------"
        Timer -start
        
        [string[]]$Filter_main += $filter
        [string[]]$Filter_main += 'WindowsApps', 'WpSystem', 'DeliveryOptimization', 'XboxGames', 'Program Files', 'Program Files (x86)', 'Windows', 'Users', 'OneCommander', '$RECYCLE.BIN', 'System Volume Information'
        
        if (($directory.Count -eq 1)) {
            $single = $true
            Shout "Single folder detected. Switching to single mode." -color Yellow
        }
        
        foreach ($i in $directory) {
            if (Test-Path -Path $i) {
                Shout "Selected folder: $i" -color Green
                if ($single) {
                    $folders += Get-Item -LiteralPath "$i" -ErrorAction SilentlyContinue
                } else {
                    $folders += Get-ChildItem -LiteralPath "$i" -Directory -Depth $apply_depth -ErrorAction SilentlyContinue
                }
            } else {
                Shout "Folder `'$i`' does not exist" -color Red
                continue
            }
        }
        
        if ($folders.Count -eq 0) {
            Shout "Folders not found or inaccessible" -color Red
            return
        }
        
        if ($filter) {
            Shout "Your filter list:" -new -color Yellow
            Shout "-----------------"
            foreach ($i in $filter){ Shout "$i" -color Yellow }
        }
        
        Shout "Processing folders:" -color Cyan -new
        Shout "$($folders -join "`n")" -color DarkBlue
        
        $primaryType = if ($priority -eq 'ico') { 'ico' } else { 'exe' }
        $secondaryType = if ($priority -eq 'ico') { 'exe' } else { 'ico' }
        Shout "Priority '$primaryType'" -new
        foreach ($folder in $folders) {
            if ($remove) {
                try {
                    if ($single) {
                        $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -ErrorAction SilentlyContinue
                    } else {
                        $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse -Depth 1 -ErrorAction SilentlyContinue
                    }
                    $desktopINI | Remove-Item -Force
                } catch {
                    Shout 'Access to the path is denied. Cant proseed with desktop.ini file. Skiping...' -color Red
                    Shout "$($folder.FullName)"
                    continue
                }
                continue
            }
            
            $shouldSkip = Test-ForbiddenFolder -Path $folder.FullName -ForbiddenFolders $Filter_main
            
            if (-not $shouldSkip) {
                Shout "Processing folder: `'$($folder.FullName)`'" -color Cyan -new
                $Files = ''
                $folder.Attributes = 'Directory', 'ReadOnly'
                [string]$full_path_folder = $folder.FullName
                $LastDirName = Split-Path -Path "$full_path_folder" -Leaf
                if (($rules) -and ($rules.ContainsKey($LastDirName))) {
                    $value = $rules[$LastDirName]
                    if (Test-Path "$full_path_folder\$value") {
                        $Files = Get-ChildItem -LiteralPath "$full_path_folder" -Filter $value
                    }
                }
                
                if (-not $Files) {
                    $Files = Find-Candidates -path $full_path_folder -search_depth $search_depth -priority $primaryType
                }
                
                if (-not $Files) {
                    Shout "$primaryType files not found, switching to $secondaryType search" -color Yellow
                    $Files = Find-Candidates -path $full_path_folder -search_depth $search_depth -priority $secondaryType
                }
                
                if ($Files) {
                    #Testing path
                    try {
                        if ($single) {
                            $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -ErrorAction SilentlyContinue
                        } else {
                            $desktopINI = Get-ChildItem -LiteralPath "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse -Depth 1 -ErrorAction SilentlyContinue
                        }
                    } catch {
                        Shout 'Access to the path is denied. Cant proseed with desktop.ini file. Skiping...' -color Red
                        Shout "$($folder.FullName)"
                        continue
                    }
                    
                    if (!($NoForce)){
                        #Forcing desktop.ini deletion
                        $desktopINI | Remove-Item -Force
                    } else {
                        if (!($desktopINI)) {
                            Shout "desktop.ini not found. Proceeding with creation" -color Green
                        } else {
                            $found = $false
                            $content = Get-Content -LiteralPath "$($desktopINI.FullName)" -ErrorAction Stop
                            foreach ($line in $content) {
                                if ($line -match '^IconResource=') {
                                    $found = $true
                                    break
                                }
                            }
                            if ($found) {
                                Shout "desktop.ini already exist. Skipping due to -NoForce flag" -color Yellow
                                continue
                            }
                        }
                    }
                    
                    #### Creating desktop.ini file starts
                    $first_part = ''
                    
                    if (($Files.DirectoryName -ne $full_path_folder)) {
                        $exe_array = ($Files.DirectoryName).Split('\')
                        $folder_array = ($full_path_folder).Split('\')
                        $diff = (Compare-Object -ReferenceObject $exe_array -DifferenceObject $folder_array).InputObject
                        foreach ($k in $diff) {
                            $first_part = $first_part + '\' + $k
                        }
                    }
                    
                    $tmpDir = (Join-Path -Path "$env:TEMP" -ChildPath ([IO.Path]::GetRandomFileName()))
                    $null = mkdir -Path $tmpDir -Force
                    $tmp = "$tmpDir\desktop.ini"
                    
                    if ($first_part) {
                        $value = '.' + "$first_part\$Files" + ',0'
                    } else {
                        $value = '.\' + $Files + ',0'
                    }
                    
                    $ini = @(
                        '[.ShellClassInfo]'
                        "IconResource=$value"
                        #"InfoTip=$exeFiles"
                        '[ViewState]'
                        'Mode='
                        'Vid='
                        'FolderType=Generic') -join "`n"
                    
                    $null = New-Item -Path "$tmp" -Value $ini
                    
                    (Get-Item -LiteralPath $tmp).Attributes = 'Archive, System, Hidden'
                    
                    $shell = New-Object -ComObject Shell.Application
                    $shell.NameSpace($full_path_folder).MoveHere($tmp, 0x0004 + 0x0010 + 0x0400)
                    #### Creating desktop.ini file ends
                    
                    Remove-Item -Path "$tmpDir" -Force
                    
                    Shout "$($Files.Name) --> $($folder.Name)" -color Green
                } else {
                    Shout "Proper file not found" -color Red
                    $foldersError += $full_path_folder
                }
            } else {
                Shout "Skipping filtered folder: $($folder.FullName)" -color DarkGray -new
            }
        }
        
        if ($remove) {
            Shout "Icons have been removed from folders! Explorer must be restarted!" -color Yellow
            Shout "Press R to restart" -color Magenta
            Shout "Press any other key to cancel restart" -color Cyan
            $key = [System.Console]::ReadKey($true)
            
            if ($key.Key -eq 'R') {
                Shout "Restarting Explorer..." -color Yellow
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                Start-Process explorer.exe
                Shout "Explorer restarted." -color Green
            } else {
                Shout "Operation cancelled." -color DarkGray
            }
        }
        
        if ($foldersError) {
            Shout "Folders with errors:" -color Red -new
            Shout "$($foldersError -join "`n")" -color Red
            Shout "Proper files not found. Try to increase search depth with -search_depth (current $search_depth)" -color Red -new
            $foldersError = @()
        }
        Shout "------------`n    DONE`n------------" -color Green
    } catch {
        Shout "$_" -color Red -new
        Shout "$($_.ScriptStackTrace)" -color Red -new -after
    }
    
    Timer -end
    
    if ($pause){
        pause
    }
}