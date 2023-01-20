#
# -----------------------------------------------------------------------------------------------
#
function local:cleanup ([String] $loc) {
	#
	if ([System.IO.File]::Exists($loc) -eq $true) {
		#
		Write-Debug @"
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
[String] $local:tmpImage = [System.IO.Path]::GetTempFileName()
#
[String] $global:tenant = "iptrack"
[String] $global:site = "RESTAPI"
#
[Reflection.Assembly] $local:ass = [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#
[Double] $local:width = 600.0
[Double] $local:height = 600.0
#
[System.Drawing.Color] $local:white = [System.Drawing.Color]::FromName('White')
[System.Drawing.SolidBrush] $local:brush = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('White'))
[System.Drawing.SolidBrush] $local:pen = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('Black'))
[System.Drawing.Font] $local:font = New-Object -Typename System.Drawing.Font "Courier New", 14
#
[String] $local:url = @"
https://$tenant.sharepoint.com/sites/$site/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/items?`$select=Material,Attachments,AttachmentFiles&`$expand=AttachmentFiles&`$top=1000000&`$filter=Attachments%20ne%200
"@
#
[String] $local:token = $(Get-Content C:\Users\Bernhard\Desktop\token.txt)
#
[Object] $local:json = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $token";"Accept"="application/json"} -Uri "$url")

foreach($tmp in $json.value) {
	#
	[String] $local:pattern = "^(.*?)\.(bmp|BMP|gif|GIF|png|PNG|jpg|JPG|jpeg|JPEG)$"
	#
	Write-Host "Material:  $($tmp.Material) 	ID: $($tmp.AttachmentFiles[0].ServerRelativeUrl  -replace '^/sites/RESTAPI/Lists/ProductObligations/Attachments/([0-9]{3,5})/.*?$','$1')"
	#
	foreach($file in $tmp.AttachmentFiles) {
		#
		if ($file.FileName -match $local:pattern) {
			#
			[Object] $local:prop = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $token";"Accept"="application/json"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($file.ServerRelativeUrl)')/Properties")
			#
			Write-Host  @"
file:	$($file.FileName) 
	written on: .... $($local:prop.vti_x005f_timelastwritten)
	size: .......... $($local:prop.vti_x005f_filesize) bytes
"@
			#
			Invoke-WebRequest -Method Get -Headers @{"Authorization"="Bearer $token";"Accept"="image/jpeg, image/png, image/gif, image/pjpeg"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($file.ServerRelativeUrl)')/`$value" -OutFile $tmpImage	# "$($file.FileName)"
			#
			[System.Drawing.Bitmap] $local:rawImg = [System.Drawing.Bitmap]::FromFile( $tmpImage, $true)
			#
			if($rawImg.PropertyIdList.Contains(274)) {
				#
				[System.Drawing.RotateFlipType] $local:rot = [System.Drawing.RotateFlipType]::RotateNoneFlipNone
				#
				switch ([BitConverter]::ToUInt16($rawImg.GetPropertyItem(274).value, 0))
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
				$local:rawImg.RotateFlip($local:rot)
				#
			}
			#
			[Double] $local:scale = [Math]::Min($local:width / $($local:rawImg.Width), $local:height / $($local:rawImg.Height))
			#
			Write-Host "scale: $local:scale rotation: $local:rot"
			#
			[Int] $local:scaleWitdh = [Convert]::ToInt32($($local:rawImg.Width) * $local:scale)
			[Int] $local:scaleHeight = [Convert]::ToInt32($($local:rawImg.Height) * $local:scale)
			#
			[System.Drawing.Bitmap] $local:scaledBitmap = New-Object -TypeName System.Drawing.Bitmap @([Convert]::ToInt32($local:width), [Convert]::ToInt32($local:height))
			#
			[System.Drawing.Graphics] $local:graph = [System.Drawing.Graphics]::FromImage($local:scaledBitmap)
			#
			$local:graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::High
			$local:graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
			$local:graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
			$local:graph.FillRectangle($local:brush, $(New-Object -TypeName System.Drawing.Rectangle @(0, 0, [Convert]::ToInt32($local:width), [Convert]::ToInt32($local:height))))
			$local:graph.DrawImage($local:rawImg, $(New-Object -TypeName System.Drawing.Rectangle @([Convert]::ToInt32(0.5 * ($local:width - $local:scaleWitdh)) , [Convert]::ToInt32(0.5 * ($local:height -  $local:scaleHeight)), $local:scaleWitdh, $local:scaleHeight)))
			$local:graph.DrawString("$($tmp.Material)   $($local:prop.vti_x005f_timelastwritten)", $local:font, $local:pen, 4, 4)
			#
			[String] $local:out = "$pwd\pics\$($($file.FileName) -replace $local:pattern,'$1').png"
			#
			$local:scaledBitmap.Save($local:out, [System.Drawing.Imaging.ImageFormat]::Png)
			#
			$local:graph.Dispose()
			#
			$local:rawImg.Dispose()
			#
			$local:scaledBitmap.Dispose()
			#
		}
		#
	}
	#
}
#
cleanup -loc $script:tmpImage
#