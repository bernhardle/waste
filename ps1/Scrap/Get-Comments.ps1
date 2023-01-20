[String] $script:tenant = 'iptrack'
#
[String] $local:list = 'Product-Master-Data'
#
[String] $script:site = 'RESTAPI'
#
[String] $script:psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'
#
[String] $local:act = $(Get-Content $(. OpenFileDialog -title 'Select input file for reading bearer token...' -type 'txt' -defpath $pwd))
#
[Object] $local:types = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -Uri "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/ContentTypes")
#
[String] $local:uri = "https://iptrack.sharepoint.com/sites/$script:site/_api/web/Lists/GetByTitle('$local:list')/items?`$select=ID,Material,Description1,ContentTypeId&`$top=10000&`$orderby=ID"
#
[Xml] $local:res = $(Invoke-WebRequest -Method GET -Headers @{'Authorization'="Bearer $local:act";'Accept'='application/xml'}  -Uri $local:uri).Content
#
[int] $local:cnt = 0
#
[int] $local:all = $res.feed.entry.content.properties.length
#
[System.Text.StringBuilder] $local:buf = [System.Text.StringBuilder]::new()
#
[void] $local:buf.AppendLine("List;Type;Material;Description;Date;Author;Comment")
#
foreach($pos in $res.feed.entry.content.properties) {
	#
	Write-Progress -Activity "Scanning $local:all items ..." -PercentComplete $([int]($local:cnt++ / $local:all * 100))
	#
	[int] $local:key = $pos.Id[0].'#text'
	#
	# $($types.content.properties | Select-Object -Property StringId,Name,Description | Where-Object { $_.StringId -eq $($pos.ContentTypeId)}).Name
	#
	[Object] $local:cmm = $(Invoke-RestMethod -Method Get -Headers @{'Authorization'="Bearer $local:act";'Accept'='application/json'} -Uri "https://iptrack.sharepoint.com/sites/$script:site/_api/web/Lists/GetByTitle('$local:list')/items($local:key)/GetComments()")
	#
	# Write-Host "$local:key $($local:cmm.value.length)"
	#
	if($local:cmm.value.length -gt 0) {
		#
		# Write-Host "Comment found at index: $local:key"
		#
		foreach($tmp in $local:cmm.value) {
			#
			[void] $local:buf.AppendLine(@"
$local:list; $($($types.content.properties | Select-Object -Property StringId,Name,Description | Where-Object { $_.StringId -eq $($pos.ContentTypeId)}).Name); $($pos.Material); $($pos.Description1); $($($local:tmp.createdDate).substring(0,10)); $($local:tmp.author.email); $($([Xml] $('<root>' + $($local:tmp.text) + '</root>')).root)
"@)
			#
		}
		#
	}
	#
}
#
ConvertFrom-Csv -Delimiter ";" -InputObject $local:buf.ToString() | Out-GridView -Title "Comments in $local:list" -PassThru |Export-Csv -Delimiter ";" -Encoding UTF8 -Path $(. SaveFileDialog -defpath $pwd -defname "comments.csv") -NoTypeInformation
#