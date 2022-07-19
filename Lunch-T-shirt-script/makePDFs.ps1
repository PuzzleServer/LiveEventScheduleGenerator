# Input file is expected to be tab-separated with no first line for header names
#    New line format must be Windows-style \r\n aka CRLF
#    Column order is TEAM NAME -> PLAYER NAME -> PLAYER EMAIL -> LUNCH OPTION
#    PLAYER NAME and LUNCH OPTION may be blank, though in the future names may not be
# Make sure the following files are all located in the same directory as this script:
#    BouncyCastle.Crypto.dll
#    Common.Logging.Core.dll
#    Common.Logging.dll
#    itext.html2pdf.dll
#    itext.io.dll
#    itext.kernel.dll
#    itext.layout.dll
#    itext.styledxmlparser.dll

# ---- CHANGE THESE VARIABLES AS NECESSARY ----

$inputFile = "E:\OneDrive\Puzzle Day\lunch-script\input.txt"
$outputFile = "E:\Desktop\lunch\index.pdf"
[String[]]$lunchList = "beet-sandwich", "border-wrap", "chicken-caesar", "garden-salad", "mushroom-wrap", "roast-beef", "roasted-turkey", "salmon-wrap", "smoked-ham", "turkey-cobb", "no-food", "none-of-the-above", "remote-participant", "Unknown"

# ---- DON'T CHANGE ANYTHING BELOW HERE ----

# Exit early to avoid clobbering existing output or erroring on missing input
if (![System.IO.File]::Exists($inputFile)) {
	Write-Output "Input file does not exist"
	Exit
	
}
if ([System.IO.File]::Exists($outputFile)) {
	Write-Output "Output file already exists"
	Exit
	
}

# Initialize needed variables
Add-Type -Path "$PSScriptRoot\itext.styledxmlparser.dll"
Add-Type -Path "$PSScriptRoot\itext.html2pdf.dll"
Add-Type -Path "$PSScriptRoot\BouncyCastle.Crypto.dll"
Add-Type -Path "$PSScriptRoot\Common.Logging.dll"
$outputDirectory = $outputFile.Substring(0, $outputFile.LastIndexOf('\'))
[System.IO.StreamReader]$inputFile = [System.IO.File]::OpenText($inputFile)
$null = [System.IO.Directory]::CreateDirectory($outputDirectory)
$null = [System.IO.Directory]::CreateDirectory("$outputDirectory\TEMP")

# Loop through each team in the input file
$fullList = $inputFile.ReadToEnd().Split([string]"`r`n", [System.StringSplitOptions]::RemoveEmptyEntries)
$inputFile.Close()
$thisTeam = ""
$thisPlayer = ""
$thisLunch = "Unknown"
$numPlayers = 0
[System.Collections.Generic.Dictionary[string, int]]$lunchCount = [System.Collections.Generic.Dictionary[string, int]]::new()
$playerTableHeader = "<table style='margin-left:auto;margin-right:auto;border-style:solid;border-width:1px;border-color:black;border-collapse:collapse'><thead><tr><th style='font-weight:bold;text-align:left;border-style:solid;border-width:1px;border-color:black'>Player</th><th style='font-weight:bold;text-align:left;border-style:solid;border-width:1px;border-color:black'>Lunch Order</th></tr></thead><tbody>"
$playerTable = $playerTableHeader
foreach ($lunch in $lunchList) {
	$lunchCount.Add($lunch, 0)
}
for ($lineNum = 0; $lineNum -lt $fullList.length; $lineNum++) {
	$line = $fullList[$lineNum].Split([char]"`t", [System.StringSplitOptions]::RemoveEmptyEntries)
	$thisTeam = $line[0]
	$numPlayers++
	
	# Update the lunch count and player table for this team
	# If the name is missing, using the email instead
	# This can later be removed once names can't be missing
	if ($line[1].length -gt 0) {
		$thisPlayer = $line[1]
	}
	else {
		$thisPlayer = $line[2]
	}
	if ($line[3].length -gt 0) {
		$thisLunch = $line[3]
	}
	$lunchCount[$thisLunch] = $lunchCount[$thisLunch] + 1
	$sanitizedPlayer = $thisPlayer.Replace([string]"&", [string]"&amp;").Replace([string]"<", [string]"&lt;").Replace([string]">", [string]"&gt;").Replace([string]"`"", [string]"&quot;").Replace([string]"'", [string]"&apos;")
	$playerTable += "<tr><td style='border-style:solid;border-width:1px;border-color:black;padding:3px'>$sanitizedPlayer</td><td style='border-style:solid;border-width:1px;border-color:black;padding:3px'>$thisLunch</td></tr>"
	
	# Update progress
	if ((([float]$lineNum / $fullList.length) -lt 0.1) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.1)) {
		Write-Output "10% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.2) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.2)) {
		Write-Output "20% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.3) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.3)) {
		Write-Output "30% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.4) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.4)) {
		Write-Output "40% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.5) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.5)) {
		Write-Output "50% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.6) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.6)) {
		Write-Output "60% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.7) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.7)) {
		Write-Output "70% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.8) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.8)) {
		Write-Output "80% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 0.9) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 0.9)) {
		Write-Output "90% complete processing teams"
	}
	elseif ((([float]$lineNum / $fullList.length) -lt 1.0) -and ((([float]$lineNum + 1.0) / $fullList.length) -ge 1.0)) {
		Write-Output "100% complete processing teams"
	}
	
	# If this is the last member of a team, output the final list for that team
	if (($lineNum + 1 -eq $fullList.length) -or ($fullList[$lineNum + 1].Split([char]"`t", [System.StringSplitOptions]::None)[0].Trim() -ne $thisTeam)) {
		$playerTable += "</tbody></table>"
		$sanitizedTeam = $thisTeam.Replace([string]"&", [string]"&amp;").Replace([string]"<", [string]"&lt;").Replace([string]">", [string]"&gt;").Replace([string]"`"", [string]"&quot;").Replace([string]"'", [string]"&apos;")
		$teamFile = "<html><head><title>$sanitizedTeam</title><style>font-family:Verdana</style><meta charset='UTF-8'></head><body><h1 style='text-align:center'>$sanitizedTeam</h1><br>$playerTable<br><h3 style='text-align:center'>Player count: $numPlayers</h3><br><table style='margin-left:auto;margin-right:auto;border-style:solid;border-width:1px;border-color:black;border-collapse:collapse'><thead><tr><th style='font-weight:bold;text-align:left;border-style:solid;border-width:1px;border-color:black'>Lunch Order</th><th style='font-weight:bold;text-align:left;border-style:solid;border-width:1px;border-color:black'>Count</th></tr></thead><tbody>"
		foreach ($lunch in $lunchList) {
			$teamFile += "<tr><td style='border-style:solid;border-width:1px;border-color:black;padding:3px'>$lunch</td><td style='border-style:solid;border-width:1px;border-color:black;padding:3px'>$($lunchCount[$lunch])</td></tr>"
			$lunchCount[$lunch] = 0
		}
		$teamFile += "</tbody></table></body></html>"
		
		# Properly encode Unicode characters
		# .NET internally handles everything as Unicode16 internally, so only surrogate pairs need to be manually combined
		# Also ignore non-visual modifiers (other ignores may need to be added to in the future)
		[System.Text.StringBuilder]$builder = [System.Text.StringBuilder]::new()
		$teamArray = $teamFile.ToCharArray()
		for ($char =  0; $char -lt $teamArray.length; $char++) {
			if ([int]($teamArray[$char]) -gt 0x7F) {
				if (([int]($teamArray[$char]) -ge 0xD800) -and ([int]($teamArray[$char]) -le 0xDFFF)) {
					if (($char + 1 -lt $teamArray.length) -and ([int]($teamArray[$char + 1]) -ge 0xD800) -and ([int]($teamArray[$char + 1]) -le 0xDFFF)) {
						$left = (([int]($teamArray[$char])) -band 0x3FF) -shl 10
						$right = ([int]($teamArray[$char + 1])) -band 0x3FF
						$null = $builder.Append("&#$($left + $right + 0x10000);")
					}
				}
				elseif (([int]($teamArray[$char]) -lt 0xFE00) -or ([int]($teamArray[$char]) -gt 0xFE0F)) {
					$null = $builder.Append("&#$([int]($teamArray[$char]));")
				}
			}
			else {
				$null = $builder.Append($teamArray[$char])
			}
		}
		
		# Create the PDF file from the generated HTML
		$filenameTeam = $thisTeam.Replace([char]"\", [char]" ").Replace([char]"/", [char]" ").Replace([char]":", [char]" ").Replace([char]"*", [char]" ").Replace([char]"?", [char]" ").Replace([char]"`"", [char]" ").Replace([char]"<", [char]" ").Replace([char]">", [char]" ").Replace([char]"|", [char]" ").Trim()
		[iText.Kernel.Pdf.PdfWriter]$writer = [iText.Kernel.Pdf.PdfWriter]::new("$outputDirectory\TEMP\$filenameTeam.pdf")
		[iText.Kernel.Pdf.PdfDocument]$pdf = [iText.Kernel.Pdf.PdfDocument]::new($writer)
		$pdf.setDefaultPageSize([iText.Kernel.Geom.PageSize]::LETTER);
		[iText.Html2pdf.ConverterProperties]$properties = [iText.Html2pdf.ConverterProperties]::new()
		$null = $properties.SetFontProvider([iText.Html2pdf.Resolver.Font.DefaultFontProvider]::new($true, $true, $true))
		[iText.Layout.Document]$document = [iText.Html2pdf.HtmlConverter]::ConvertToDocument($builder.ToString(), $pdf, $properties)
		$document.Close()
		
		# Reset variables for next iteration
		$numPlayers = 0
		$playerTable = $playerTableHeader
	}
	$thisTeam = ""
	$thisPlayer = ""
	$thisLunch = "Unknown"
}

# Merge the PDfs and output the final file
$tempDir = Get-ChildItem "$outputDirectory\TEMP" | Sort-Object -Property name
[iText.Kernel.Pdf.PdfWriter]$finalWriter = [iText.Kernel.Pdf.PdfWriter]::new("$outputFile")
[iText.Kernel.Pdf.PdfDocument]$finalPdf = [iText.Kernel.Pdf.PdfDocument]::new($finalWriter)
foreach ($file in $tempDir) {
	[iText.Kernel.Pdf.PdfReader]$thisReader = [iText.Kernel.Pdf.PdfReader]::new($file.fullname)
	[iText.Kernel.Pdf.PdfDocument]$thisPdf = [iText.Kernel.Pdf.PdfDocument]::new($thisReader)
	[System.Collections.Generic.List[int]]$pageList = [System.Collections.Generic.List[int]]::new()
	for ($i = 1; $i -le $thisPdf.GetNumberOfPages(); $i++) {
		$pageList.Add($i)
	}
	$null = $thisPdf.CopyPagesTo($pageList, $finalPdf)
	$thisPdf.Close()
}
$finalPdf.Close()
Remove-Item "$outputDirectory\TEMP" -Recurse
