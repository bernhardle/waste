[String] $local:list = 'Product-Master-Data'

[String] $local:sit = 'RESTAPI'

[String] $script:psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'

[String] $local:act = $(Get-Content C:\Users\Bernhard\Desktop\act.txt)

[Object] $local:csv = Import-Csv -Header @('Material', 'Kurztext', 'Items') -Encoding 'utf8' -Delimiter ';' -Path $(. $script:psmodules/OpenFileDialog -title 'Select input file for reading ...' -type 'csv' -defpath $pwd)

$local:csv | Out-GridView  -Title "Add-ListEntry" -Wait

[System.Collections.Hashtable] $local:hdr = @{'Authorization'="Bearer $local:act"; 'Accept'='application/json;odata=verbose'; 'Content-Type'='application/json'; 'If-Match'='*'; 'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Site $local:sit -AccessToken $act)"}

foreach($pos in 1..$csv.length) {

	[String] $local:mat = $csv[$pos].Material
	
	if ($local:mat.length -ge 5) {

		[String] $local:uri = "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items?`$select=Id,Material&`$filter=Material eq '$local:mat'"
		
		# Write-Host $local:uri
		
		[System.Xml.XmlElement] $local:ime = $(Invoke-RestMethod -Method GET -Headers @{'Authorization'="Bearer $local:act";'Accept'='application/xml'}  -Uri $local:uri)
		
		if ($ime.content.properties.Material -eq $local:mat) {
		
			Write-Host "ERROR: Material '$local:mat' already existent - skipping ..."
		
		} else {
		
			[String] $local:json = "{""ContentTypeId"": ""0x01003FAF714C6769BF4FA1B36DCF47ED659702009768DCF9E8E5424EB2E68CAB20681245"",""Material"":""$($($csv[$pos]).Material)"", ""Description1"":""$($($csv[$pos]).Kurztext)"",""ItemListId"":[$($($csv[$pos]).Items)]}"
		
			Write-Host "POST: $local:mat $local:json"
			
			Invoke-RestMethod -Method POST -Headers $local:hdr -Body $local:json -ContentType 'application/json' -Uri "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items"
		
		}
	}
}

return $null