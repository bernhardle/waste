#
[String] $script:tenant = 'iptrack'
#
# -----------------------------------------------------------------------------------------------
#
function local:ConvertTo-Base64Url {
     param([byte[]]$Bytes)
     [Convert]::ToBase64String($Bytes).
         TrimEnd("=").
         Replace("+", "-").
         Replace("/", "_")
}
#
# -----------------------------------------------------------------------------------------------
#
function local:getAccessToken {
	#
	#	Siehe auch: 	https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread
	#					https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials
	#
	[String] $private:realm='801ebad3-0ef0-432b-9be5-90593a424825'			# iptrack.onmicrosoft.com
	#
	[String] $private:clientId = 'c5d349bd-f59d-40b6-9baa-7772c5215b3d'		# Waste-Bulk-Insert-Helper
	#
	[String] $local:thumbprint = 'F9E92DBEBE48D1C0B19089C7C28F82A1C2688993'	
	#
	[System.Security.Cryptography.X509Certificates.X509Certificate2] $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $local:thumbprint}
	#
	$header = @{
		alg = "RS256" 
		typ = "JWT"
		x5t = $(ConvertTo-Base64Url $cert.GetCertHash())
	}
	#
	$payload = @{
		aud = "https://login.microsoftonline.com/$private:realm/oauth2/v2.0/token"
		iss = $private:clientId
		sub = $private:clientId
		jti = $($(New-Guid).ToString())
		nbf = $([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
		exp = $([DateTimeOffset]::UtcNow.AddMinutes(10).ToUnixTimeSeconds())
	}
	#
	$headerJson  = $header  | ConvertTo-Json -Compress
	$payloadJson = $payload | ConvertTo-Json -Compress
	#
	$headerEncoded = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($headerJson))
	$payloadEncoded = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payloadJson))
	#
	$dataToSign = "$headerEncoded.$payloadEncoded"
	#
	$signatureEncoded = ConvertTo-Base64Url $([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert).SignData([Text.Encoding]::UTF8.GetBytes($dataToSign), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1))
	#
	$jwt = "$dataToSign.$signatureEncoded"
	#
	$body = @{
		client_id             = $private:clientId
		scope                 = "https://$script:tenant.sharepoint.com/.default"
		grant_type            = "client_credentials"
		client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
		client_assertion      = $jwt
	}
	#
	[Object] $private:result = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$private:realm/oauth2/v2.0/token" -body $body
	#
	return $private:result.access_token
	#
}
#
# -----------------------------------------------------------------------------------------------
#
#
[String] $local:list = 'Product-Master-Data'
#
[String] $local:sit = 'RESTAPI'
#
[String] $script:psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'
#
[String] $local:act = local:getAccessToken		# $(Get-Content C:\Users\Bernhard\Desktop\act.txt)
#
[Object] $local:csv = Import-Csv -Header @('Material', 'Kurztext', 'Items') -Encoding 'utf8' -Delimiter ';' -Path $(. $script:psmodules/OpenFileDialog -title 'Select input file for reading ...' -type 'csv' -defpath $pwd)
#
$local:csv | Out-GridView  -Title "Add-ListEntry" -Wait
#
[System.Collections.Hashtable] $local:hdr = @{'Authorization'="Bearer $local:act"; 'Accept'='application/json;odata=verbose'; 'Content-Type'='application/json'; 'If-Match'='*'; 'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Site $local:sit -AccessToken $act)"}
#
foreach($pos in 1..$csv.length) {
	#
	[String] $local:mat = $csv[$pos].Material
	#
	if ($local:mat.length -ge 5) {
		#
		[String] $local:uri = "https://$script:tenant.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items?`$select=Id,Material&`$filter=Material eq '$local:mat'"
		#
		# Write-Host $local:uri
		#
		[System.Xml.XmlElement] $local:ime = $(Invoke-RestMethod -Method GET -Headers @{'Authorization'="Bearer $local:act";'Accept'='application/xml'}  -Uri $local:uri)
		#
		if ($ime.content.properties.Material -eq $local:mat) {
			#
			Write-Host "ERROR: Material '$local:mat' already existent - skipping ..."
			#
		} else {
			#
			[String] $local:json = "{""ContentTypeId"": ""0x01003FAF714C6769BF4FA1B36DCF47ED659702009768DCF9E8E5424EB2E68CAB20681245"",""Material"":""$($($csv[$pos]).Material)"", ""Description1"":""$($($csv[$pos]).Kurztext)"",""ItemListId"":[$($($csv[$pos]).Items)]}"
			#
			Write-Host "POST: $local:mat $local:json"
			#
			Invoke-RestMethod -Method POST -Headers $local:hdr -Body $local:json -ContentType 'application/json' -Uri "https://$script:tenant.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$list')/items"
			#
		}
		#
	}
	#
}
#
return $null