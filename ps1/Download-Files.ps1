#
# -----------------------------------------------------------------------------------------------
#
function local:cleanup ([String] $loc) {
	#
	if ([System.IO.File]::Exists($loc) -eq $true) {
		#
		Write-Verbose @"
SAP-DR-Reporting.ps1::cleanup (...)	Deleting temporary file '$loc'.
"@
		#
		[System.IO.File]::Delete($loc)
		#
	}
}
#
# -----------------------------------------------------------------------------------------------
#
function local:decorate ([String] $inFile, [String] $outFile, [String] $headingL = "Links oben", [String] $headingR = "Rechts oben", [String] $subheading = "") {
	#
	[System.Drawing.Bitmap] $private:rawImg = [System.Drawing.Bitmap]::FromFile($inFile, $true)
	#
	if($private:rawImg.PropertyIdList.Contains(274)) {
		#
		[System.Drawing.RotateFlipType] $private:rotation = [System.Drawing.RotateFlipType]::RotateNoneFlipNone
		#
		switch ([BitConverter]::ToUInt16($rawImg.GetPropertyItem(274).value, 0))
		{
			#
			2	{
				$private:rotation = [System.Drawing.RotateFlipType]::RotateNoneFlipX
			}
			#
			3	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate180FlipNone 
			}
			#
			4	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate180FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			5	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate90FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			6	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate90FlipNone
			}
			#
			7	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate270FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			8	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate270FlipNone
			}
			#
		}
		#
		$private:rawImg.RotateFlip($private:rotation)
		#
	}
	#
	[Double] $private:scale = [Math]::Min($script:decoWidth / $($private:rawImg.Width), $script:decoHeight / $($private:rawImg.Height))
	#
	Write-Verbose @"
SAP-DR-Reporting.ps1::decorate (): Scale: $private:scale rotation: $private:rotation"
"@
	#
	[Int] $private:scaleWitdh = [Convert]::ToInt32($($private:rawImg.Width) * $private:scale)
	[Int] $private:scaleHeight = [Convert]::ToInt32($($private:rawImg.Height) * $private:scale)
	#
	[System.Drawing.Bitmap] $private:scaledBitmap = New-Object -TypeName System.Drawing.Bitmap @([Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))
	#
	[System.Drawing.Graphics] $private:graph = [System.Drawing.Graphics]::FromImage($private:scaledBitmap)
	#
	$private:graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::High
	$private:graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
	$private:graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$private:graph.FillRectangle($script:decobrush, $(New-Object -TypeName System.Drawing.Rectangle @(0, 0, [Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))))
	$private:graph.DrawImage($private:rawImg, $(New-Object -TypeName System.Drawing.Rectangle @([Convert]::ToInt32(0.5 * ($script:decoWidth - $private:scaleWitdh)) , [Convert]::ToInt32(0.5 * ($script:decoHeight -  $private:scaleHeight)), $private:scaleWitdh, $private:scaleHeight)))
	#
	$private:graph.DrawString($headingL, $script:decoFontBold, $script:decoPen, $script:decoRectMat)
	$private:graph.DrawString($headingR, $script:decoFontBold, $script:decoPen, $script:decoRectDat)
	$private:graph.DrawString($subheading, $script:decoFontNormal, $script:decoPen, $script:decoRectItm)
	#
	$private:scaledBitmap.Save($outFile, [System.Drawing.Imaging.ImageFormat]::Png)
	#
	$private:graph.Dispose()
	#
	$private:rawImg.Dispose()
	#
	$private:scaledBitmap.Dispose()
	#
}
#
[String] $script:decotmpImage = [System.IO.Path]::GetTempFileName()
#
[String] $global:tenant = "iptrack"
[String] $global:site = "RESTAPI"
#
[Reflection.Assembly] $local:ass = [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#
[Double] $script:decoWidth = 1500.0
[Double] $script:decoHeight = 1000.0
#
# [System.Drawing.Color] $script:decoWhite = [System.Drawing.Color]::FromName('White')
[System.Drawing.SolidBrush] $script:decobrush = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('White'))
[System.Drawing.SolidBrush] $script:decopen = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('Black'))
[System.Drawing.Font] $script:decofontNormal = New-Object -Typename System.Drawing.Font "Courier New", $(12 * $script:decoWidth * 0.001)
[System.Drawing.FontStyle] $script:decostyleBold = [System.Drawing.Fontstyle]::Bold
[System.Drawing.Font] $script:decofontBold = New-Object -Typename System.Drawing.Font $script:decofontNormal, $script:decostyleBold
[System.Drawing.RectangleF] $script:decorectMat = New-Object -TypeName System.Drawing.RectangleF $($script:decoWidth * 0.003), $($script:decoWidth * 0.004), $($script:decoWidth*0.795), $($script:decoWidth * 0.020)
[System.Drawing.RectangleF] $script:decorectDat = New-Object -TypeName System.Drawing.RectangleF  $($script:decoWidth * 0.800), $($script:decoWidth * 0.004), $($script:decoWidth*1.000), $($script:decoWidth * 0.020)
[System.Drawing.RectangleF] $script:decorectItm = New-Object -TypeName System.Drawing.RectangleF  $($script:decoWidth * 0.003), $($script:decoWidth * 0.025), $($script:decoWidth*1.000), $($script:decoWidth * 0.020)
#
[String] $local:act = $(Get-Content C:\Users\Bernhard\Desktop\token.txt)
#
foreach ($row in $list) {
	#
	foreach ($decoRowVal in $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/json"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/Items($($row.Item_Id))/AttachmentFiles").value) {
		#
		[Object] $local:pro = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/json"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/Properties")
		#
		[String] $private:pattern = "^(.*?)\.(bmp|BMP|gif|GIF|png|PNG|jpg|JPG|jpeg|JPEG)$"
		#
		if ($decoRowVal.FileName -match $private:pattern) {
			#
			[String] $private:outFile = "$pwd\pics\$($row.Base_Material)_$($row.Item_Typ)_$($row.Item_Id)_$($($decoRowVal.FileName) -replace $private:pattern,'$1').png"
			#
			Write-Host  @"
			
SAP-DR-Reporting.ps1::main (): Downloading image '$($decoRowVal.FileName)' 
|	device: ...... $($local:pro.wic_x005f_System_x005f_Photo_x005f_CameraManufacturer) $($local:pro.wic_x005f_System_x005f_Photo_x005f_CameraModel)
|	created: ..... $($local:pro.vti_x005f_timelastwritten)
|	disk size: ... $($local:pro.vti_x005f_filesize) bytes
|	writing to: . $outfile
"@
			#
			Invoke-WebRequest -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="image/jpeg, image/png, image/gif, image/pjpeg"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/`$value" -OutFile $local:tmpImage
			#
			if ($($row.Item_Id) -ne $($row.Base_Id)) {
				. decorate -inFile $private:tmpImage -outFile $private:outFile -headingL "$($row.Base_Material) $($row.Base_Kurztext)" -headingR $($local:prop.vti_x005f_timelastwritten) -subheading "$($row.Item_Typ)($($row.Item_material)): $($row.Item_Kurztext)"
			} else {
				. decorate -inFile $private:tmpImage -outFile $private:outFile -headingL "$($row.Base_Material) $($row.Base_Kurztext)" -headingR $($local:prop.vti_x005f_timelastwritten)
			}
			#
		} else {
			#
			[String] $private:outFile = "$pwd\pics\$($row.Base_Material)_$($row.Item_Typ)_$($row.Item_Id)_$($decoRowVal.FileName)"
			#
			Write-Host  @"
			
SAP-DR-Reporting.ps1::main (): Downloading file '$($decoRowVal.FileName)'
|	created: ..... $($local:pro.vti_x005f_timelastwritten)
|	disk size: ... $($local:pro.vti_x005f_filesize) bytes
|	writing to: . $outfile
"@
			#
			Invoke-WebRequest -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/octet-stream"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/`$value" -OutFile $outFile
			#
		}
	}
}
#
cleanup -loc $private:tmpImage
#