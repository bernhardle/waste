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
[String] $local:mode = 'Replace-List' # 'Add-ListEntry'
#
[String] $local:list = 'Product-Master-Data'
#
[String] $local:field = 'ItemListId' # 'Duty_x002d_ListId' # 'ItemListId' #   'Part_x002d_ListId' 
#
[String] $local:sit = 'RESTAPI'
#
[String] $script:psmodules = 'C:\Users\Bernhard\Documents\XML Formulare\Common\ps1\'
#
[String] $local:act = local:getAccessToken		# $(Get-Content C:\Users\Bernhard\Desktop\act.txt)
#
[Object] $local:csv = Import-Csv -Header @('Material', $field) -Encoding 'utf8' -Delimiter ';' -Path $(. OpenFileDialog -title 'Select input file for reading ...' -type 'csv' -defpath $pwd)
#
$local:csv | Out-GridView  -Title "$local:mode" -Wait
#
[System.Collections.Hashtable] $local:hdr = @{'Authorization'="Bearer $local:act"; 'Accept'='application/json;odata=verbose'; 'Content-Type'='application/json'; 'If-Match'='*'; 'X-RequestDigest'="$(. "$psmodules\Get-SPOFormDigestValue.ps1" -Site $local:sit -AccessToken $local:act)"}
#
foreach($pos in 1..$csv.length) {

	[String] $local:mat = $csv[$pos].Material
	
	[Int] $local:entry = $($($csv[$pos]).$field)
	
	if ($local:mat.length -ge 5) {

		[String] $local:uri = "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$local:list')/items?`$select=Id,$field&`$filter=Material eq '$local:mat'"
		
		[System.Xml.XmlElement] $local:ime = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/xml"}  -Uri "$uri")

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
			Write-Host -ForegroundColor Red "ERROR: Entry '$entry' already present in list '$field' for material '$mat' - skipping ..."
			#
		} else {
			#
			Write-Host -ForegroundColor Green "PATCH: Material: $mat Feld: $field Wert: $local:json"
			#
			Invoke-RestMethod -Method PATCH -Headers $local:hdr -Body $local:json -ContentType 'application/json' -Uri "https://iptrack.sharepoint.com/sites/$local:sit/_api/web/Lists/GetByTitle('$local:list')/items($local:idx)"
			#
		}
	}
}
#
return $null
#