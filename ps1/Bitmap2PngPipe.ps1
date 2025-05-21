#	___________________________________________________________________
#
#	(c) Bernhard Schupp, Frankfurt (2010-2025)
#
#	Version:
#		2010-11-13:	Erstellt.
#		2010-11-26:	Gueltigkeitsbereiche der Variablen geaendert.
#		2010-12-09:	Meldungen angepasst.
#		2010-12-12:	Meldungen aus PS cmdlets.
#		2016-06-29:	.emf und vollstaendige Dateinamen.
#       2020-06-04: Graustufenumwandlung als weitere Option.
#
param([Boolean] $greyscale = $false, [String] $format = "png", [Boolean] $verbose = $true, [Boolean] $debug = $false)
#
#	Argumente:
#       -greyscale
#       -format
#		-verbose
#		-debug
#
begin
{
    #
    Add-Type -AssemblyName System.Drawing
	#
	#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
	#	und entsprechend den Parametern $verbose und $debug neu setzen.
	#
	[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
	[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
	#
	if($verbose -eq $true) {
		$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
	}
	if($debug -eq $true) {
		$DebugPreference = [System.Management.Automation.ActionPreference]::Continue
	}
    #
    [System.Drawing.SolidBrush] $script:decoBrush = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('White'))
    #
	Write-Verbose @"
Bitmap2PngPipe.ps1: 'greyscale' ... $greyscale.
Bitmap2PngPipe.ps1: 'format' ...... $format
Bitmap2PngPipe.ps1: 'verbose' ..... $verbose.
Bitmap2PngPipe.ps1: 'debug' ....... $debug.
"@
    #
    [String] $local:add = "_resized"
    #
    switch($format) {
        #
        "png" {
            [System.Drawing.Imaging.ImageFormat] $script:fmt = [System.Drawing.Imaging.ImageFormat]::Png
            [String] $script:postfix = $("$local:add.png")
        }
        #
        "bmp" {
            [System.Drawing.Imaging.ImageFormat] $script:fmt = [System.Drawing.Imaging.ImageFormat]::Bmp
            [String] $script:postfix = $("$local:add.bmp")
        }
        #
        "gif" {
            [System.Drawing.Imaging.ImageFormat] $script:fmt = [System.Drawing.Imaging.ImageFormat]::Gif
            [String] $script:postfix = $("$local:add.gif")
        }
        "jpg" {
            [System.Drawing.Imaging.ImageFormat] $script:fmt = [System.Drawing.Imaging.ImageFormat]::Jpg
            [String] $script:postfix = $("$local:add.jpg")
        }
        default {
			Write-Warning @"
Bitmap2PngPipe.ps1::begin: Format '$format' unknown. Exiting ...
"@
            exit 2
        }
    }
    #
    [Double] $script:decoWidth = 1560.0
    [Double] $script:decoHeight = 1040.0
    #
}

process
{
	[String] $local:inp =  $($_.FullName)
	[String] $local:oup = "$($inp -replace '\.emf|\.bmp$|\.jpeg$|\.jpg$|\.gif$', $script:postfix)"
	#
	Write-Debug @"

Bitmap2PngPipe.ps1::process: Loading ... $inp
Bitmap2PngPipe.ps1::process: Writing ... $oup
"@
	#
	[System.Drawing.Bitmap] $local:rawImg = [System.Drawing.Bitmap]::FromFile($inp)
    #
    Write-Verbose @"

Bitmap2PngPipe.ps1::process: Input image width .... $($local:rawImg.Width)
Bitmap2PngPipe.ps1::process: Input image height ... $($local:rawImg.Height)
"@
    #
    [System.Drawing.RotateFlipType] $local:rot = [System.Drawing.RotateFlipType]::Rotate270FlipNone
    #
    if($local:rawImg.PropertyIdList.Contains(274)) {
        #
        [System.Drawing.RotateFlipType] $local:rot = [System.Drawing.RotateFlipType]::RotateNoneFlipNone
        #
        switch ([BitConverter]::ToUInt16($local:rawImg.GetPropertyItem(274).value, 0))
        {
            #
            2	{
                $local:rot = [System.Drawing.RotateFlipType]::RotateNoneFlipX
            }
            #
            3	{
                $local:rot = [System.Drawing.RotateFlipType]::Rotate180FlipNone 
            }
            #
            4	{
                $local:rot = $([System.Drawing.RotateFlipType]::Rotate180FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
            }
            #
            5	{
                $local:rot = $([System.Drawing.RotateFlipType]::Rotate90FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
            }
            #
            6	{
                $local:rot = [System.Drawing.RotateFlipType]::Rotate90FlipNone
            }
            #
            7	{
                $local:rot = $([System.Drawing.RotateFlipType]::Rotate270FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
            }
            #
            8	{
                $local:rot = [System.Drawing.RotateFlipType]::Rotate270FlipNone
            }
            #
        }
        #
        # $local:rawImg.RotateFlip($local:rot)
        #
    }
    #
    [Double] $private:scale = [Math]::Min($script:decoWidth / $($private:rawImg.Width), $script:decoHeight / $($private:rawImg.Height))
    #
    [Double] $private:scaleWitdh = $($local:rawImg.Width) * $local:scale
    [Double] $private:scaleHeight = $($local:rawImg.Height) * $local:scale
    #
    Write-Host @"
Bitmap2PngPipe.ps1::process: scale ...... $local:scale
Bitmap2PngPipe.ps1::process: width ...... $private:scaleWitdh
Bitmap2PngPipe.ps1::process: height ..... $private:scaleHeight
Bitmap2PngPipe.ps1::process: rotation ... $local:rot
Bitmap2PngPipe.ps1::process: xpos ....... $([Convert]::ToInt32(0.5 * ($script:decoWidth - $private:scaleWitdh)))
Bitmap2PngPipe.ps1::process: ypos ....... $([Convert]::ToInt32(0.5 * ($script:decoHeight -  $private:scaleHeight)))
"@
    #
    [System.Drawing.Bitmap] $private:scaledBitmap = New-Object -TypeName System.Drawing.Bitmap @([Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))
    #
    [System.Drawing.Graphics] $private:graph = [System.Drawing.Graphics]::FromImage($private:scaledBitmap)
    #
	$private:graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::High
	$private:graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
	$private:graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$private:graph.FillRectangle($script:decoBrush, $(New-Object -TypeName System.Drawing.Rectangle @(0, 0, [Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))))
	$private:graph.DrawImage($local:rawImg, $(New-Object -TypeName System.Drawing.Rectangle @([Convert]::ToInt32(0.5 * ($script:decoWidth - $private:scaleWitdh)), [Convert]::ToInt32(0.5 * ($script:decoHeight -  $private:scaleHeight)), [Convert]::ToInt32($private:scaleWitdh), [Convert]::ToInt32($private:scaleHeight))))
    #
    if($greyscale -eq $true) {
    	Write-Debug @"
Bitmap2PngPipe.ps1::process: Converting to greyscale ... 
"@
        #
        for([int] $local:ypos = 0 ; $local:ypos -lt $($local:scaledBitmap.Height) ; $local:ypos = $($local:ypos + 1)) {
            #
            for([int] $local:xpos = 0 ; $local:xpos -lt $($local:scaledBitmap.Width) ; $local:xpos = $($local:xpos + 1)) {
                [System.Drawing.Color] $local:pixcol = $local:scaledBitmap.GetPixel($local:xpos, $local:ypos)
                [int] $local:tmp = $(0.3 * $local:pixcol.R + 0.59 * $local:pixcol.G + 0.11 * $local:pixcol.B)
                [System.Drawing.Color] $local:pixgry = [System.Drawing.Color]::FromArgb($local:tmp, $local:tmp, $local:tmp)
                $local:scaledBitmap.SetPixel($local:xpos, $local:ypos, $local:pixgry)
            }
            #
        }
        #
    }
    #
	$local:scaledBitmap.Save($oup, [System.Drawing.Imaging.ImageFormat]::Png)
    #
	$private:graph.Dispose()
	#
	$private:rawImg.Dispose()
	#
	$private:scaledBitmap.Dispose()
}

end
{

	Write-Debug @"
Bitmap2PngPipe.ps1::end: End of pipeline.
"@
	#
	#	Werte vor Start des Skripts wiederherstellen
	#
	$VerbosePreference = $script:saveVerbosePref
	$DebugPreference = $script:saveDebugPref
	#
}