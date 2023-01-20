
[String] $local:result = $(& 'C:\Program Files\curl\CURL.EXE' -i -H 'Authorization: Bearer' -s 'https://iptrack.sharepoint.com/_vti_bin/client.svc')

[String] $local:bearer_realm=$($result -replace '.*WWW-Authenticate: Bearer realm="([0-9a-z-]*)",client_id="([0-9a-z-]*)".*','$1')

# Alternativ kann für den Tenant "iptrack" für $local:bearer_realm der feste Wert "801ebad3-0ef0-432b-9be5-90593a424825" verwendet werden.

[String] $local:app_id=$($result -replace '.*WWW-Authenticate: Bearer realm="([0-9a-z-]*)",client_id="([0-9a-z-]*)".*','$2')

# Alternativ kann für den Tenant "iptrack" für $local:app_id der feste Wert "00000003-0000-0ff1-ce00-000000000000" verwendet werden.

[String] $local:grant_type="grant_type=client_credentials"

[String] $local:cl_id="client_id=15c1da3a-5a13-4adb-b288-d8957afd13e4@$bearer_realm"

[String] $local:cl_secret="client_secret=e6omj241geUo0UY0gSkrLcp+IFwCJkdG8wyhv0rmeKg="

[String] $local:res="resource=$app_id/iptrack.sharepoint.com@$bearer_realm"

[String] $local:url="https://accounts.accesscontrol.windows.net/$bearer_realm/tokens/OAuth/2"

[String] $local:content_type="Content-Type: application/x-www-form-urlencoded"

[String] $local:access_token=$($(& 'C:\Program Files\curl\CURL.EXE' -X POST -H $content_type --data-urlencode $grant_type --data-urlencode $cl_id --data-urlencode $cl_secret --data-urlencode $res -s $url) -replace '{("\w*":".*",)*(?:"\w*":"(.*)")}','$2')

[String] $local:ear_liste=$(& 'C:\Program Files\curl\CURL.EXE' -X GET -H "Authorization: Bearer $access_token" -H "Accept: application/xml" -s "https://iptrack.sharepoint.com/sites/RESTAPI/_api/web/Lists/GetByTitle('Einfachste')/items")

echo $ear_liste

# [String] $local:tmp = $(& 'C:\Program Files\curl\CURL.EXE' -X GET -H "Authorization: Bearer $access_token" -H "Accept: application/xml" -s "https://iptrack.sharepoint.com/sites/Kanzlei/_api/web/lists(guid'0396a2a1-1387-45ce-8ae8-8223bc1745bc')/Items(1)?`$select=Title,Attachments,Attachmentfiles&`$expand=AttachmentFiles")

# [String] $local:listUri="https://iptrack.sharepoint.com/sites/Kanzlei/_api/web/Lists/GetByTitle('EAR-Stammdaten')/items?`$select=Title,Gewicht,Kategorie"
# [System.Collections.Hashtable] $local:headerMap =@{"Authorization"="Bearer $access_token";"Accept"="application/xml"}
# Invoke-RestMethod -Method Get -Headers $headerMap -Uri "$($listUri)?`$select=Title,Gewicht,unit,Kategorie" -OutFile invoke.xml