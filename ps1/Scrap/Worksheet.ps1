#	Listeneinträge mit PNPOnline erstellen
#	Powershell 7.1 installieren
#	Update-Module -Name PnP.PowerShell als Admin
#	Connect-PNPOnline -Url https://iptrack.sharepoint.com/sites/RESTAPI
#
#	Alternativ kann man das Access Token speichern und
#	und in die Umgebung des REST weiterleiten, es hat
#	bessere Berechtigungen als die in SP gehosteten APP
#
Get-PnPAppAuthAccessToken | Out-File -FilePath "$env:tmp\token.txt"
#
[String] $act = Get-Content -Path "$env:tmp\token.txt"
#
#	-------------------------------------------------------------------------------------------------------------------------------
[String] $psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'
#
[String] $tnt = 'iptrack'
#
[String] $sit = 'RESTAPI'
#
[String] $act = $(. "$psmodules\Get-SPOAccessToken.ps1")
#
[String] $gid = 'b7ffe2cb-00de-400d-8ee1-cdd124ef222c' # guid von 'Einfachste Liste'
#
[String] $uri = "https://$tnt.sharepoint.com/sites/$sit/_api/web/Lists(guid'$gid')"
#
#	Die nächste Zeile liefert den Type des Listenelements
#
[String] $typ = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -Uri "$uri`?`$select=ListItemEntityTypeFullName").entry.content.properties.ListItemEntityTypeFullName
#
[System.Collections.Hashtable]  $bdy = @{'__metadata' = @{'type' = $typ};'Titel' = 'Hallo' ; 'ContentTypeId' = 'Element'}
#
[System.Collections.Hashtable] $local:hdr = @{
'Authorization'="Bearer $act";
'Accept'='application/json;odata=verbose';
'Content-Type'='application/json';
'If-Match'='*';
'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Site $sit -AccessToken $act)"}
#
Invoke-RestMethod -Method POST -Headers $hdr -Body $bdy -ContentType 'application/json' -Uri "$uri/items"		 # Failed 'missing authorization'
#
#	Delete Operation
#
[System.Collections.Hashtable] $local:hdr = @{
'Authorization'="Bearer $act";
'Accept'='application/json;odata=verbose';
'Content-Type'='application/json';
'If-Match'='*';
'X-HTTP-Method'='DELETE'}
#
[String] $uri = "https://iptrack.sharepoint.com/sites/RESTAPI/_api/web/Lists(guid'$gid')"
#
Invoke-RestMethod -Method DELETE -Headers $hdr -ContentType 'application/json' -Uri "$uri/items(_)"
#
