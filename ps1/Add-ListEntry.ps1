[String] $local:mode = 'Replace-List' # 'Add-ListEntry'

[String] $local:list = 'Product-Master-Data'

[String] $local:field = 'ItemListId' # 'Duty_x002d_ListId' # 'ItemListId' #   'Part_x002d_ListId' 

[String] $local:sit = 'RESTAPI'

[String] $script:psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'

[String] $local:act = $(Get-Content C:\Users\Bernhard\Desktop\act.txt)

[Object] $local:csv = Import-Csv -Header @('Material', $field) -Encoding 'utf8' -Delimiter ';' -Path $(. OpenFileDialog -title 'Select input file for reading ...' -type 'csv' -defpath $pwd)

$local:csv | Out-GridView  -Title "$local:mode" -Wait

[System.Collections.Hashtable] $local:hdr = @{'Authorization'="Bearer $local:act"; 'Accept'='application/json;odata=verbose'; 'Content-Type'='application/json'; 'If-Match'='*'; 'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Site $local:sit -AccessToken $act)"}

foreach($pos in 1..$csv.length) {

	[String] $local:mat = $csv[$pos].Material
	
	[Int] $local:entry = $($($csv[$pos]).$field)
	
	if ($local:mat.length -ge 5) {

		[String] $local:uri = "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items?`$select=Id,$field&`$filter=Material eq '$local:mat'"
		
		[System.Xml.XmlElement] $local:ime = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"}  -Uri "$uri")

		[Int] $local:idx = $($ime.content.properties.Id[0].'#text')
		
		# Write-Host "Material # $($mat.padLeft(12)) Index $($idx.padLeft(5)) $field $($ime.content.properties.$field.element) Neu: $entry"
		
		[Boolean] $local:skip = $false
		
		[String] $local:json = "{""$field"":["
		
		switch ($local:mode) {
			#
			'Add-ListEntry' {
				#
				foreach($key in $ime.content.properties.$field.element) {
					#
					if ($key -eq $local:entry) {
						#
						$local:skip = $true
						#
					}
					#
					$local:json = $local:json + $key + ','
					#
				}
				#
				$local:json = $local:json + "$($($csv[$pos]).$field)]}"
				#
				break ;
			}
			#
			'Replace-List' {
				#
				$local:json = $local:json + $local:entry + "]}"
				#
				break ;
			}
			#
			default {
				#
				break ;
				#
			}
		}
		#
		if ($local:skip)	{
			#
			Write-Host "ERROR: Entry '$entry' already present in list '$field' for material '$mat' - skipping ..."
			#
		} else {
			#
			Write-Host "PATCH: Material: $mat Feld: $field Wert: $local:json"
			#
			Invoke-RestMethod -Method PATCH -Headers $local:hdr -Body $local:json -ContentType 'application/json' -Uri "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items($local:idx)"
			#
		}
	}
}

return $null