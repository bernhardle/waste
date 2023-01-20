#
#	Retrieves all items from Product-Obligations repository (list)
#
param([String] $psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\', [String] $AccessToken = $(throw "Access token missing."), [String] $FilePath = '')
#
function local:getPOListItems([String] $act, [String] $out) {
	#
	[String] $local:salespackage = '0x01003FAF714C6769BF4FA1B36DCF47ED65970400F6E244CC367A26439227A016D1DF32CD'
	[String] $local:product = '0x01003FAF714C6769BF4FA1B36DCF47ED659702009768DCF9E8E5424EB2E68CAB20681245'
	[String] $local:device = '0x01003FAF714C6769BF4FA1B36DCF47ED65970300A3EE85F8979D4A4ABFF77A6A7F09ED9D'
	[String] $local:electric = '0x01003FAF714C6769BF4FA1B36DCF47ED6597030100F1F1D05B8F05D14082CB1F3E0D35D4E3'
	[String] $local:scippart = '0x0100F1AB00CC48EDDB41BA1F7EFF656DCD3A0100296B386716B0164C8B8E5C744D07CAA8'
	#
	[String] $local:uri = "https://iptrack.sharepoint.com/sites/RESTAPI/_api/web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/items`?`$select=Id,Material,Description1,ItemListId`&`$filter=ContentTypeId eq '$salespackage'"
	#
	[System.Collections.Hashtable] $local:hdr = @{"Authorization"="Bearer $act";"Accept"="application/xml"}
	#
	Invoke-RestMethod -Method Get -Headers $hdr -Uri "$($uri)" -OutFile $out
	#
}
#
if ($FilePath.length -eq 0) {
	#
	$FilePath = $(. $psmodules\SaveFileDialog.ps1 -title 'Select output file for writing XML converted DSD package data' -defpath $pwd -delete $true -type 'xml')
	#
	if ($FilePath.length -eq 0) {
		Write-Verbose @"
Operation aborted by user.
"@
		exit 1
	}
	#
}
#
. getPOListItems -act $AccessToken -out $FilePath
#