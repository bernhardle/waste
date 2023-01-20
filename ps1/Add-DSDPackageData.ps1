#
param([String] $psmodules = "C:\Users\Bernhard\Eigene Dateien\XML Formulare\Common\ps1", [String] $Tenant = 'iptrack', [String] $Site = 'RESTAPI', [String] $ListGUID = '5cd75f66-486f-49fa-8176-b3e74fc8a10d', [String] $AccessToken)
#
function local:addDSDPackageData([String] $inp = $(throw "Add-DSDPackageData.ps1::convertDSDPackageData(...) Input file path missing."), [String] $act = $(throw "Add-DSDPackageData.ps1::convertDSDPackageData(...) Access token missing.")) {
	#
	[System.Collections.Hashtable] $local:hdr = @{
		'Authorization'="Bearer $act";
		'Accept'='application/json;odata=verbose';
		'Content-Type'='application/json';
		'If-Match'='*';
		'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Tenant $Tenant -Site $Site -AccessToken $act)"}
	#
	#	[System.Collections.Hashtable] $bdy = @{'ContentTypeId'='0x01003FAF714C6769BF4FA1B36DCF47ED65970400F6E244CC367A26439227A016D1DF32CD'; 'Material'='Test-010292'; 'Description1'="Testverpackung aus REST API PUT"; 'Duty_x002d_ListId'=@(3,4); 'VV_x002d_Alu'='125'; 'VV_x002d_Steel'='0'; 'Paper'='2341';'VV_x002d_Plastic'='265';'VV_x002d_Tinplate'='87'}
	#
	#	'0x01003FAF714C6769BF4FA1B36DCF47ED65970400F6E244CC367A26439227A016D1DF32CD'	# ContentTypeId der Verkaufsverpackung
	#
	#	Invoke-RestMethod -Method POST -Headers $hdr -Body $(ConvertTo-JSON -InputObject $bdy) -ContentType 'application/json' -Uri "https://$Tenant.sharepoint.com/sites/$Site/_api/web/lists(guid'$ListGUID')/items"
	#
	[String] $private:json = @"
{
    "odata.metadata": "https://iptrack.sharepoint.com/sites/RESTAPI/_api/$metadata#SP.ListData.ProductObligationsListItems/@Element",
    "odata.type": "SP.Data.ProductObligationsListItem",
    "odata.id": "a3c47d88-a460-4cc8-a72c-869237e54132",
    "odata.etag": "\"1\"",
    "odata.editLink": "Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/Items(569)",
    "FileSystemObjectType": 0,
    "Id": 569,
    "ServerRedirectedEmbedUri": null,
    "ServerRedirectedEmbedUrl": "",
    "ID": 569,
    "ContentTypeId": "0x01003FAF714C6769BF4FA1B36DCF47ED65970400F6E244CC367A26439227A016D1DF32CD",
    "Title": null,
    "Modified": "2021-02-03T20:16:53Z",
    "Created": "2021-02-03T20:16:53Z",
    "AuthorId": 1073741822,
    "EditorId": 1073741822,
    "OData__UIVersionString": "1.0",
    "Attachments": false,
    "GUID": "515594bb-259f-4c73-b964-7cae73b1d996",
    "ComplianceAssetId": null,
    "Material": "Test-322",
    "Description1": "Testverpackung #2",
    "Weight": null,
    "ItemListId": [],
    "VV_x002d_Alu": 1025,
    "VV_x002d_Steel": 254,
    "Paper": 841,
    "VV_x002d_Plastic": 265,
    "VV_x002d_Tinplate": 187,
    "Duty_x002d_ListId": [
        3,
        4
    ],
    "Part_x002d_ListId": [],
    "Pieces": null
}
"@
	#
	# [String] $local:typ = 'Sales-Packaging'
	#
	# [String] $local:typ = 'Electric-Device'
	#
	[String] $local:typ = 'Product'
	#
	# [String[]] $hdr = @("material","packaging","description","alu","ppk","tin","plastics")
	#
	# [String[]] $hdr = @('material', 'weight', 'description', 'be', 'bg', 'ch', 'g1', 'g2', 'g3', 'cy', 'cz', 'de', 'ee', 'fi', 'fr', 'gr', 'nl')
	#
	[String[]] $hdr = @('material', 'weee', 'batt', 'tvvv' , 'description')
	#
	$csv = Import-Csv -Header $hdr -Encoding 'utf8' -Delimiter ';' -Path $inp
	#
	foreach($pos in 1..$csv.length) {	#$csv.length
		#
		# [Object[]] $dus = @($csv.be[$pos], $csv.bg[$pos], $csv.ch[$pos], $csv.g1[$pos], $csv.g2[$pos], $csv.g3[$pos], $csv.cy[$pos], $csv.cz[$pos], $csv.de[$pos], $csv.ee[$pos], $csv.fi[$pos], $csv.fr[$pos], $csv.gr[$pos], $csv.nl[$pos])
		#
		# $val = @{'Material'="$($csv.packaging[$pos])"; 'Description1'="$($csv.description[$pos])"; 'Duty_x002d_List'='3'; 'VV_x002d_Alu'="$($csv.alu[$pos])"; 'VV_x002d_Steel'='0'; 'Paper'="$($csv.ppk[$pos])";'VV_x002d_Plastic'="$($csv.plastics[$pos])";'VV_x002d_Tinplate'="$($csv.tin[$pos])"}
		#
		# $val = @{'Material'="$($csv.material[$pos])"; 'Description1'="$($csv.description[$pos])"; 'Weight'=$("$($csv.weight[$pos])" -replace '^([0-9]+)(?:,([0-9]+))$','$1.$2'); 'Duty_x002d_List'=$dus}
		#
		[String[]] $ims = @('','','')
		#
		[Int] $local:i = 0
		#
		if ($($csv.weee[$pos]).length -gt 0) {
			$ims[$local:i++] = $csv.weee[$pos]
		}
		if ($($csv.batt[$pos]).length -gt 0) {
			$ims[$local:i++] = $csv.batt[$pos]
		}
		if ($($csv.tvvv[$pos]).length -gt 0) {
			$ims[$local:i++] = $csv.tvvv[$pos]
		}
		#
		$val = @{'Material'="$($csv.material[$pos])"; 'Description1'="$($csv.description[$pos])"; 'ItemList'=$ims[0..$($local:i - 1)]}
		#
		if ($val['Material']) {
			# 
			$val | ConvertTo-JSON | Write-Host
			#
			Add-PnPListItem -List 'Product-Master-Data' -ContentType $local:typ -Values $val
			#
		}
		#
	}
	#
}
[String] $local:InputPath = $(. OpenFileDialog -title 'Select input file for reading DSD package data' -type 'csv' -defpath $pwd)
#
if ([System.IO.File]::Exists($local:InputPath) -eq $true) {
	#
	. addDSDPackageData -inp $local:InputPath -act $AccessToken
	#
}
Write-Host -Message @"
Bye bye ...
"@
