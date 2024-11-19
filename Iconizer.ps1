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
			Write-Host "$color doesn't exists in System.ConsoleColor" -ForegroundColor Red
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
		if (($OpenFileDialog.ShowDialog() -eq 'OK')) {
			$file = $OpenFileDialog.FileName
		}
	} else {
		$OpenFolderDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
		$OpenFolderDialog.Description = 'Select a folder'
		$OpenFolderDialog.Rootfolder = 'MyComputer'
		$OpenFolderDialog.ShowNewFolderButton = $false
		if ($OpenFolderDialog.ShowDialog($tempForm) -eq 'OK') {
			$directory = $OpenFolderDialog.SelectedPath
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

function pull {
	[CmdletBinding()]
	param (
		[Alias('d')]
		[string[]]$directory,
		[Alias('f')]
		[ValidateSet('ico', 'bmp', 'png', 'jpg')]
		$format = 'ico',
		[Alias('i')]
		[int]$index = 0,
		[Alias('dep')]
		[int]$depth = 0,
		[Alias('s')]
		[switch]$small,
		[Alias('l')]
		[string]$log
	)

	$TypeDefinition = @'
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Collections.Generic;
using System.Drawing.Drawing2D;

public static class ImagingHelper
{
    public static bool ConvertToIcon(Bitmap inputBitmap, Stream output, int[] sizes)
    {
        if (inputBitmap == null || sizes == null || sizes.Length == 0)
            return false;

        List<MemoryStream> imageStreams = new List<MemoryStream>();
        foreach (int size in sizes)
        {
            Bitmap newBitmap = ResizeImage(inputBitmap, size, size);
            if (newBitmap == null)
                return false;
            MemoryStream memoryStream = new MemoryStream();
            newBitmap.Save(memoryStream, ImageFormat.Png);
            imageStreams.Add(memoryStream);
        }

        BinaryWriter iconWriter = new BinaryWriter(output);
        if (output == null || iconWriter == null)
            return false;

        int offset = 0;

        iconWriter.Write((byte)0);
        iconWriter.Write((byte)0);

        iconWriter.Write((short)1);

        iconWriter.Write((short)sizes.Length);

        offset += 6 + (16 * sizes.Length);

        for (int i = 0; i < sizes.Length; i++)
        {
            iconWriter.Write((byte)sizes[i]);
            iconWriter.Write((byte)sizes[i]);
            iconWriter.Write((byte)0);
            iconWriter.Write((byte)0);
            iconWriter.Write((short)0);
            iconWriter.Write((short)32);
            iconWriter.Write((int)imageStreams[i].Length);
            iconWriter.Write((int)offset);
            offset += (int)imageStreams[i].Length;
        }

        for (int i = 0; i < sizes.Length; i++)
        {
            iconWriter.Write(imageStreams[i].ToArray());
            imageStreams[i].Close();
        }

        iconWriter.Flush();

        return true;
    }

    public static bool ConvertToIcon(Stream input, Stream output, int[] sizes)
    {
        Bitmap inputBitmap = (Bitmap)Bitmap.FromStream(input);
        return ConvertToIcon(inputBitmap, output, sizes);
    }

    public static bool ConvertToIcon(string inputPath, string outputPath, int[] sizes)
    {
        using (FileStream inputStream = new FileStream(inputPath, FileMode.Open))
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(inputStream, outputStream, sizes);
        }
    }

    public static bool ConvertToIcon(Image inputImage, string outputPath, int[] sizes)
    {
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(new Bitmap(inputImage), outputStream, sizes);
        }
    }

    public static Bitmap ResizeImage(Image image, int width, int height)
    {
        var destRect = new Rectangle(0, 0, width, height);
        var destImage = new Bitmap(width, height);

        destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

        using (var graphics = Graphics.FromImage(destImage))
        {
            graphics.CompositingMode = CompositingMode.SourceCopy;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            using (var wrapMode = new ImageAttributes())
            {
                wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
            }
        }

        return destImage;
    }
}
'@
		try {
			Add-Type -TypeDefinition $TypeDefinition -ReferencedAssemblies 'System.Drawing'
			$MemberDefinition = @(
				'[DllImport("Shell32.dll", SetLastError=true)]'
				'public static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);'
				''
				'[DllImport("gdi32.dll", SetLastError=true)]'
				'public static extern bool DeleteObject(IntPtr hObject);'
			) -join "`n"
			Add-Type -Namespace Win32API -Name Icon -MemberDefinition $MemberDefinition
		} catch {
			Shout "Type 'Win32API.Icon' already exist. Relaunch the console. Error:`n $_"
		}

		try {
			if (!($directory)) {
				$directory = SelectPath -files
				if ($directory){
					$file_GUI = $true
				}
			}

			Shout "Let's start extracting icons from exes" -color Green -new
			Shout "-----------------"
			Timer -start

			function GetICO {
				param (
					[string]$p
				)

				$hi, $low = 0, 0
				$null = [Win32API.Icon]::ExtractIconEx($($p), $index, [ref]$hi, [ref]$low, 1)
				$handle = if ($small) { $low } else { $hi }

				try {
					$icon = [System.Drawing.Icon]::FromHandle($handle)
				} catch {
					return $false
				}

				$hi, $low, $handle | Where-Object { $_ } | ForEach-Object { [Win32API.Icon]::DeleteObject($_) } | Out-Null
				return $icon
			}

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
							$icon = GetICO -p $($_path.FullName)

							if ($icon){
								if ($format -eq 'ico') {
									$out_png = Join-Path -Path "$($_path.Directory.FullName)" -ChildPath "$($_path.BaseName)_tmp.png"
									$out_ico = Join-Path -Path "$($_path.Directory.FullName)" -ChildPath "$($_path.BaseName)_icon.ico"
									$icon.ToBitmap().Save($out_png)
									if ($small){
										$sizes = ( 16 )
									} else {
										$sizes = ( 32 )
									}

									$null = [ImagingHelper]::ConvertToIcon("$out_png", "$out_ico", $sizes)
									if (Test-Path $out_png){
										Remove-Item $out_png -Force -ErrorAction SilentlyContinue
									}
									$out = $out_ico
								} else {
									$out = Join-Path -Path $_path.Directory.FullName -ChildPath "$($_path.BaseName)_icon.$format"
									$icon.ToBitmap().Save($out)
								}
								if (Test-Path -Path "$out"){
									Shout "Icon was extracted to:`n $out" -color Green
								} else {
									Shout "Icon was NOT extracted from:`n $($_path.FullName)" -color Red
								}
							} else {
								Shout "Can't find any icon in:`n $($_path.FullName)" -color Red
							}
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
}

function apply {
	[CmdletBinding()]
	param (
		[Alias('d')]
		[string[]]$directory,
		[Alias('p')]
		[ValidateSet('any', 'icon', 'folder')]
		$priority,
		[Alias('f')]
		[string[]]$filter,
		[Alias('s')]
		[switch]$single,
		[Alias('rm')]
		[switch]$remove,
		[Alias('l')]
		[string]$log,
		[Alias('dep')]
		$dependencies
	)
	
		$ErrorActionPreference = 'Stop'
		$pattern_regex = '[^\w\s]'
		$pattern_regex_symbols = '[!#$%&()+;@^_{}~№]'
		$pattern_regex_digits = '\d+'

		try {
			if (!($directory)) {
				$directory = SelectPath
			}
			if (!($directory)) {
				return
			}
			Shout "Let's start applying icons to folders" -color Green -new
			Shout "-----------------"
			Timer -start

			foreach ($i in $directory) {
				if (Test-Path -Path $i) {
					Shout "$i" -color Green
					if ($single) {
						$folders += Get-Item -Path $i -ErrorAction SilentlyContinue
					} else {
						$folders += Get-ChildItem -Path $i -Directory -ErrorAction SilentlyContinue
					}
				} else {
					Shout "Folder `'$i`' does not exist" -color Red
					continue
				}
			}
			
			if ($remove) {
				Shout "Icons have been removed from folders! Restarting explorer!"
				Stop-Process -Name explorer -Force; Start-Process explorer
			} elseif ($filter) {
				Shout "Your filter list:" -new -color Yellow
				Shout "-----------------"
				foreach ($i in $filter){ Shout "$i" -color Yellow }
			}
			Shout "The magic begins!" -color Cyan -new
			Shout "-----------------"
			[string[]]$Filter_main += $filter
			[string[]]$Filter_main += 'WindowsApps', 'WpSystem', 'DeliveryOptimization', 'XboxGames', 'Program Files', 'Program Files (x86)', 'Windows', 'Users', 'OneCommander'
			
			foreach ($folder in $folders) {
				if ($Filter_main -notcontains $($folder.Name)) {
					try {
						Get-ChildItem -Path "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse -Depth 1 | Remove-Item -Force
						
						if ($remove){
							Stop-Process -Name explorer -Force
							Start-Process explorer
							continue
						}
						
						$array_string = @()
						$exeFiles = ''
						$exeFile_name_checked = @()
						$array_exes = @()
						
						$folder.Attributes = 'Directory', 'ReadOnly'
					} catch {
						Shout 'Access to the path is denied.' -color Red
						Shout "$($folder.FullName)"
					}
					
					[string]$name_folder = ($folder.Name).ToLower()
					$name_folder = $name_folder -replace $pattern_regex, '' -replace "$pattern_regex_digits", '' -replace "$pattern_regex_symbols", ''
					[string]$full_path_folder = $folder.FullName
					
					$LastDirName = Split-Path -Path $full_path_folder -Leaf
					if (($dependencies) -and ($dependencies.ContainsKey($LastDirName))) {
						$value = $dependencies[$LastDirName]
						if (Test-Path $full_path_folder\$value) {
							$exeFiles = Get-ChildItem -Path $full_path_folder -Filter $value
						}
					}
					
					if (!($exeFiles)) {
						$tmpr = $name_folder.Trim().Split(' ')
						foreach ($j in $tmpr) {
							$array_string += $j
						}

						foreach ($item in $array_string) {
							$files = Get-ChildItem -Path $full_path_folder -Recurse -Filter "*$item*.exe"
							Write-Verbose -Message $item
							if ($files) { $exeFile_name_checked += $item }
						}
						
						$get_first = $exeFile_name_checked[0]
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter "*$get_first*.exe" -Recurse | Select-Object -First 1
						foreach ($i in $exeFiles){ Write-Verbose $($i.ToString()) }
					}
					
					if ((!($exeFiles)) -or ($priority -eq 'folder')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter "$name_folder.ico" -Recurse | Select-Object -First 1
					} elseif ((!($exeFiles)) -or ($priority -eq 'icon')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter 'icon.ico' | Select-Object -First 1
					} elseif ((!($exeFiles)) -or ($priority -eq 'any')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter '*.ico' -Recurse | Select-Object -First 1
					}
					
					if (!($exeFiles)) {
						$exeFiles = (Get-ChildItem -Path $full_path_folder -Filter '*.exe' -Recurse)
						
						foreach ($exeFile in $exeFiles) {
							[string]$name_exe = ($exeFile.BaseName).ToLower() -replace "$pattern_regex_symbols", '' -replace "$pattern_regex", '' -replace "$pattern_regex_digits", ''
							if (($name_folder -like $name_exe) `
								-or ($name_folder -match $name_exe) `
								-or ($name_folder.Contains($name_exe)) `
								-or ($name_exe.Contains($name_folder))) {
								
								$array_exes += $exeFile
							} elseif ($name_exe) {
								$array_exes += $exeFile
							}
						}
						$exeFiles = $array_exes[0]
					}
					
					if ($exeFiles) {
						$first_part = ''
						[string]$name_exe = ($exeFiles.BaseName).ToLower()
						
						Write-Verbose -Message "exe name: $name_exe"

						if (($exeFiles.DirectoryName -ne $full_path_folder)) {
							$exe_array = ($exeFiles.DirectoryName).Split('\')
							$folder_array = ($full_path_folder).Split('\')
							$diff = (Compare-Object -ReferenceObject $exe_array -DifferenceObject $folder_array).InputObject

							foreach ($k in $diff) {
								$first_part = $first_part + '\' + $k
							}
						}

						$tmpDir = (Join-Path -Path $env:TEMP -ChildPath ([IO.Path]::GetRandomFileName()))
						$null = mkdir -Path $tmpDir -Force
						$tmp = "$tmpDir\desktop.ini"

						if ($first_part) {
							$value = '.' + "$first_part\$exeFiles" + ',0'
						} else {
							$value = '.\' + $exeFiles + ',0'
						}

						$ini = @(
							'[.ShellClassInfo]'
							"IconResource=$value"
							"InfoTip=$exeFiles"
							'[ViewState]'
							'Mode='
							'Vid='
							'FolderType=Generic') -join "`n"

						$null = New-Item -Path $tmp -Value $ini
						
						(Get-Item -Path $tmp).Attributes = 'Archive, System, Hidden'
						
						$shell = New-Object -ComObject Shell.Application
						$shell.NameSpace($full_path_folder).MoveHere($tmp, 0x0004 + 0x0010 + 0x0400)
						
						Remove-Item -Path $tmpDir -Force
						
						Shout "$($exeFiles.Name) --> $($folder.Name)" -color Green
						Write-Verbose -Message "exe and index: $value"
						
					} else {
						Shout "Proper exe not found in $full_path_folder" -color Yellow
					}
				} else {
					Shout "Folder `'$folder`' filtered" -color Yellow
				}
			}
		} catch {
			Shout "$_" -color Red -new
			Shout "$($_.ScriptStackTrace)" -color Red -new -after
		}

	Timer -end
}