#
#	(c) Bernhard Schupp, Frankfurt (2021-2024)
#
#	Version:
#		2021-01-23:	Erstellt.
#		2021-04-19:	Verarbeitung von Bundle Items im Stylesheet prexsl.
#		2021-05-01:	Vierter Menuepunkt zum Anzeigen der Details von Items
#		2021-10-05:	Menuepunkt zum Speichern des Bearer Tokens
#		2021-10-19:	Menuepunkt zum Herunterladen der Kommentare/Notizen
#		2021-11-15:	Menuepunkt zum Berechnen des Transfers Product->Sales-Packaging
#		2022-03-01: Neues Add-In installiert mit Gueltigkeit bis 01.03.2023
#		2022-06-05: Fuer Nutzung des hierarchischen Aufbaus der ContentTypeID angepasst.
#		2022-09-20: Ablaufdatum
#		2022-10-01: Bereit fuer Products-in-Lots, enthaelt alle Stylesheets des Datums
#		2022-10-16: Menuepunkt zum Herunterladen der Anhaenge
#		2022-12-22:	Extraktion Laenderliste wie initiale Version
#		2023-01-05:	Transfer auf mehrere Ziele erweitert, Kopierverzeichnis aus Parameter
#		2023-02-04:	Neues Add-In installiert mit Gueltigkeit bis 03.02.2024, hinzu [AppContext]::SetSwitch (...) für PS7+
#		2023-02-06:	Das ist nun die Vollversion des Skripts.
#		2023-02-09:	Hinweis auf Ablaufdatum des SharePoint Secrets.
#		2023-10-09: [Linux] Changed Add-Type to look by assemblyname rather than location
#		2023-10-09: [Linux] Changed PtrToStringAuto -> PtrToStringBSTR, see: https://github.com/dotnet/runtime/issues/35632#issuecomment-621507916
#		2024-02-04: Neues Add-In installiert mit Gueltigkeit bis 27.02.2025
#		2024-02-04: Branch 'pre-2024' merged into 'main'
#		2024-03-22:	Export of product duties by country and batch complete
#	Original:
#		XML Formulare/Abfallwirtschaft/ps1/SAP-DR-Reporting.ps1
#	Verweise:
#		-/-
#	Verwendet:
#		-/-
#
param([String] $DataDir = $pwd, [Boolean] $verbose = $false, [Boolean] $debug = $false, [String] $copies = $null) 
#
Clear-Host
#
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
#
#	To allow transformations including the 'document()' function.
#
[AppContext]::SetSwitch('Switch.System.Xml.AllowDefaultResolver', $true)
#
[String] $script:tenant = 'iptrack'
#
[String] $script:site = 'RESTAPI'
#
[int] $script:rows = 10000
#
[DateTime] $script:expiration = [DateTime]"2025-02-01"
#
# -----------------------------------------------------------------------------------------------
#
[TimeSpan] $local:remaining = $($script:expiration - $(Get-Date))
#
if($local:remaining.days -lt 30 -and $local:remaining.days -ge 0) {
	#
	Write-Host @"
	*********************************************************************************************

	The validity of the SharePoint secret used in this script will expire in $($remaining.days) days.

	*********************************************************************************************
	
"@
#
} elseif ($remaining.days -lt 0) {
	#
	Write-Host @"
	***********************************************************************************************

	The validity of the SharePoint secret used in this script has expired since $script:expiration.

	***********************************************************************************************

"@
	#
	Read-Host -Prompt "Press any key to abort."
	#
	exit
	#
}
#
# -----------------------------------------------------------------------------------------------
#
function local:getAccessToken ([String] $phrase) {
	#
	#	Siehe auch: 	https://docs.microsoft.com/de-de/sharepoint/dev/sp-add-ins/create-and-use-access-tokens-in-provider-hosted-high-trust-sharepoint-add-ins
	#					https://anexinet.com/blog/getting-an-access-token-for-sharepoint-online/
	#
	[String] $private:realm='801ebad3-0ef0-432b-9be5-90593a424825'
	#
	[String] $private:url="https://accounts.accesscontrol.windows.net/$realm/tokens/OAuth/2"
	#
	[String] $private:clientId = 'f33cdb15-73eb-4f13-b424-9ed9c2cab531'
	#
	[String] $private:scrambled = 
'76492d1116743f0423413b16050a5345MgB8AGkAMQByAFUAdQBKAE0AbgBJAFgAYwBnACsAQwBFAFYAZABPAGoAMQBWAHcAPQA9AHwAMwBjAGQANwA4ADYAYwA1AGQAMwA4ADcAMgBjADIAZQA0AGIAZQAyAGYANgBhADgAZABmADIAMQA1AGYAOABjAGEAMAA2ADYAYQAwAGYANgA2ADAAMQBiAGIAMABjAGQAYwAwADkANwA5ADYAOQAyAGMAMABiADQANgA5ADAANwBjADEAOABkADkAYwBiADgAMQA0AGMAMwAzAGIAYwBjADAAZAAyADQANgA5ADUAOAA0ADEAOQA0AGIAYwBkAGYAYQA5ADEAOQBkAGUAYgBlADEAOAA5ADAAOABhAGEANwBhADMAYgBhAGYAMgBiADgAYwAzADgAMwAyADgAYgBlADEANwA5AGIAZgA4ADgAZQA0ADQAZAAxADcAMABiADAAYQA5ADEAOABiAGIANwBkADAANQAyADAAZQAxAGIAYgBjADAAMwA2ADYAYgA2AGYANAAwADcANgA0AGQAYgA0AGMAZQA5AGUANAA1AGEAMwA4ADQAYQA3AGMAOAA3ADUA'
	#
	if (($phrase.length -lt 16) -or ($phrase.length -gt 32)) {
		throw "[Fatal] SAP-DR-Reporting.ps1::getAccessToken(...): Key required with length of 16...32 chars."
	}
	[System.Text.ASCIIEncoding] $local:enc = New-Object System.Text.ASCIIEncoding
	[Byte []] $private:key = $enc.GetBytes($phrase + "0" * (32 - $phrase.length))
	#
	[System.Collections.Hashtable] $local:body=@{
	grant_type='client_credentials';
	client_id="$clientId@$realm";
	client_secret="$($private:scrambled | ConvertTo-SecureString -key $key | ForEach-Object {[Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($_))})";
	resource="00000003-0000-0ff1-ce00-000000000000/$tenant.sharepoint.com@$realm"
	}
	#
	[Microsoft.PowerShell.Commands.WebResponseObject] $aut = Invoke-WebRequest -Method POST -ContentType 'application/x-www-form-urlencoded' -Body $body -Uri $url
	#
	return $($aut.Content | ConvertFrom-Json).access_token
	#
}
#
# -----------------------------------------------------------------------------------------------
#
function local:getListItems ([String] $act, [String] $mod, [String] $out) {
	#
	switch -e ($mod) 
	{	
		#
		'duties'	{
			#
			[String] $local:gid = 'e88e0881-c2ee-4cb9-9742-c228ef9ed458' 
			#
			[String] $local:uri = "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$local:gid')/items?`$select=Id,ContentTypeId,Duty,Country,Description12,Batch,Rank,Modified&`$top=$script:rows&`$orderby=ID"
			#
			$null = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -OutFile $out -Uri $local:uri)
			#
			break
			#
		}
		#
		'products'	{
			#
			[String] $local:gid = '5cd75f66-486f-49fa-8176-b3e74fc8a10d'
			#
			[String] $local:uri = "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$local:gid')/items?`$select=Id,ContentTypeId,Material,Description1,Weight,ItemListId,Reference_x002d_ProductId,VV_x002d_Alu,VV_x002d_Steel,Paper,VV_x002d_Plastic,VV_x002d_Tinplate,Duty_x002d_ListId,Part_x002d_ListId,Pieces,REACH,SCIP_x002d_Number,Attachments,Modified&`$top=$script:rows&`$order=ID"
			#
			$null = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -OutFile $out -Uri $local:uri)
			#
			break
		}
		#
		'types' 	{
			#
			[String] $local:gid = '5cd75f66-486f-49fa-8176-b3e74fc8a10d'
			#
			[String] $local:uri = "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$local:gid')/ContentTypes?`$top=$script:rows"
			#
			$null = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -OutFile $out -Uri $local:uri)
			#
			break
			#
		}
		#
		'fields' 	{
			#
			[String] $local:gid = '5cd75f66-486f-49fa-8176-b3e74fc8a10d'
			#
			[String] $local:uri = "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$local:gid')/Fields?`$select=InternalName,Title,Description,TypeAsString,TypeDisplayName,TypeShortDescription&`$top=$script:rows"
			#
			$null = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -OutFile $out -Uri $local:uri)
			#
			break
			#
		}
		#
	}
	#
}
#
# -----------------------------------------------------------------------------------------------
#
function local:getComments () {
	#
	[String] $private:gid = '5cd75f66-486f-49fa-8176-b3e74fc8a10d'
	#
	[Xml] $local:res = $(Invoke-WebRequest -Method GET -Headers @{'Authorization'="Bearer $act";'Accept'='application/xml'}  -Uri "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$private:gid')/items?`$select=ID,Material,Description1,ContentTypeId&`$top=$script:rows&`$orderby=ID").Content
	#
	[Object] $local:types = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $act";"Accept"="application/xml"} -Uri "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$private:gid')/ContentTypes")
	#
	[int] $local:cnt = 0
	#
	[int] $local:all = $res.feed.entry.content.properties.length
	#
	[System.Text.StringBuilder] $local:buf = [System.Text.StringBuilder]::new()
	#
	[void] $local:buf.AppendLine("Type;Material;Description;Date;Author;Comment")
	#
	foreach($pos in $res.feed.entry.content.properties) {
		#
		Write-Progress -Id  234 -Activity "Scanning $local:all items ..." -PercentComplete $([int]($local:cnt++ / $local:all * 100))
		#
		[int] $local:key = $pos.Id[0].'#text'
		#
		[Object] $local:cmm = $(Invoke-RestMethod -Method Get -Headers @{'Authorization'="Bearer $local:act";'Accept'='application/json'} -Uri "https://$script:tenant.sharepoint.com/sites/$script:site/_api/Web/Lists(guid'$private:gid')/items($local:key)/GetComments()")
		#
		if($local:cmm.value.length -gt 0) {
			#
			foreach($tmp in $local:cmm.value) {
				#
				[void] $local:buf.AppendLine(@"
$($($types.content.properties | Select-Object -Property StringId,Name,Description | Where-Object { $_.StringId -eq $($pos.ContentTypeId)}).Name); $($pos.Material); $($pos.Description1); $($($local:tmp.createdDate).ToShortDateString()); $($local:tmp.author.email); $($([Xml] $('<root>' + $($local:tmp.text) + '</root>')).root)
"@)
				#
			}
			#
		}
		#
	}
	#
	return $local:buf.ToString()
	#
}
#
# -----------------------------------------------------------------------------------------------
#
[Double] $script:decoWidth = 1500.0
[Double] $script:decoHeight = 1000.0
#
[System.Drawing.SolidBrush] $script:decoBrush = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('White'))
[System.Drawing.SolidBrush] $script:decoPen = New-Object -TypeName System.Drawing.SolidBrush $([System.Drawing.Color]::FromName('Black'))
[System.Drawing.Font] $script:decofontNormal = New-Object -Typename System.Drawing.Font "Courier New", $(12 * $script:decoWidth * 0.001)
[System.Drawing.FontStyle] $script:decostyleBold = [System.Drawing.Fontstyle]::Bold
[System.Drawing.Font] $script:decofontBold = New-Object -Typename System.Drawing.Font $script:decofontNormal, $script:decostyleBold
[System.Drawing.RectangleF] $script:decorectMat = New-Object -TypeName System.Drawing.RectangleF $($script:decoWidth * 0.003), $($script:decoWidth * 0.004), $($script:decoWidth*0.795), $($script:decoWidth * 0.020)
[System.Drawing.RectangleF] $script:decorectDat = New-Object -TypeName System.Drawing.RectangleF  $($script:decoWidth * 0.800), $($script:decoWidth * 0.004), $($script:decoWidth*1.000), $($script:decoWidth * 0.020)
[System.Drawing.RectangleF] $script:decorectItm = New-Object -TypeName System.Drawing.RectangleF  $($script:decoWidth * 0.003), $($script:decoWidth * 0.025), $($script:decoWidth*1.000), $($script:decoWidth * 0.020)
#
#
function local:decorate ([String] $decoInFile, [String] $decoOutFile, [String] $headingL = "[top left]", [String] $headingR = "[bottom right]", [String] $subheading = "") {
	#
	[System.Drawing.Bitmap] $private:rawImg = [System.Drawing.Bitmap]::FromFile($decoInFile, $true)
	#
	if($private:rawImg.PropertyIdList.Contains(274)) {
		#
		[System.Drawing.RotateFlipType] $private:rotation = [System.Drawing.RotateFlipType]::RotateNoneFlipNone
		#
		switch ([BitConverter]::ToUInt16($rawImg.GetPropertyItem(274).value, 0))
		{
			#
			2	{
				$private:rotation = [System.Drawing.RotateFlipType]::RotateNoneFlipX
			}
			#
			3	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate180FlipNone 
			}
			#
			4	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate180FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			5	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate90FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			6	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate90FlipNone
			}
			#
			7	{
				$private:rotation = $([System.Drawing.RotateFlipType]::Rotate270FlipNone -or [System.Drawing.RotateFlipType]::RotateNoneFlipX)
			}
			#
			8	{
				$private:rotation = [System.Drawing.RotateFlipType]::Rotate270FlipNone
			}
			#
		}
		#
		$private:rawImg.RotateFlip($private:rotation)
		#
	}
	#
	[Double] $private:scale = [Math]::Min($script:decoWidth / $($private:rawImg.Width), $script:decoHeight / $($private:rawImg.Height))
	#
	Write-Verbose @"
SAP-DR-Reporting.ps1::decorate (): Scale: $private:scale rotation: $private:rotation"
"@
	#
	[Int] $private:scaleWitdh = [Convert]::ToInt32($($private:rawImg.Width) * $private:scale)
	[Int] $private:scaleHeight = [Convert]::ToInt32($($private:rawImg.Height) * $private:scale)
	#
	[System.Drawing.Bitmap] $private:scaledBitmap = New-Object -TypeName System.Drawing.Bitmap @([Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))
	#
	[System.Drawing.Graphics] $private:graph = [System.Drawing.Graphics]::FromImage($private:scaledBitmap)
	#
	$private:graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::High
	$private:graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
	$private:graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$private:graph.FillRectangle($script:decoBrush, $(New-Object -TypeName System.Drawing.Rectangle @(0, 0, [Convert]::ToInt32($script:decoWidth), [Convert]::ToInt32($script:decoHeight))))
	$private:graph.DrawImage($private:rawImg, $(New-Object -TypeName System.Drawing.Rectangle @([Convert]::ToInt32(0.5 * ($script:decoWidth - $private:scaleWitdh)) , [Convert]::ToInt32(0.5 * ($script:decoHeight -  $private:scaleHeight)), $private:scaleWitdh, $private:scaleHeight)))
	#
	$private:graph.DrawString($headingL, $script:decoFontBold, $script:decoPen, $script:decoRectMat)
	$private:graph.DrawString($headingR, $script:decoFontBold, $script:decoPen, $script:decoRectDat)
	$private:graph.DrawString($subheading, $script:decoFontNormal, $script:decoPen, $script:decoRectItm)
	#
	$private:scaledBitmap.Save($decoOutFile, [System.Drawing.Imaging.ImageFormat]::Png)
	#
	$private:graph.Dispose()
	#
	$private:rawImg.Dispose()
	#
	$private:scaledBitmap.Dispose()
	#
}
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
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/SAP-DR-Recycling-Countries.xslt (2022-12-22)
#
[Xml] $script:ccsxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	xmlns:xls="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:p="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:x="urn:schemas-microsoft-com:office:excel" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	exclude-result-prefixes="com dty p ss x xls">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:key name="ccx" use="xls:Cell[7]/xls:Data" match="xls:Row" />
	<xsl:key name="ccc" use="Property [@Name='Empfangsland']" match="Object" />
	<xsl:param name="SAP-DR-Recycling-Countries.MasterData" select="''" />
	<xsl:variable name="master" select="document (`$SAP-DR-Recycling-Countries.MasterData)/p:root" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="blockWidth" select="90" />
	<xsl:template match="*">
		<xsl:message>
			<xsl:value-of select="concat (substring (`$constHyp140, 1, `$blockWidth), '&#xA;')"/>
			<xsl:text>
 *** ERROR: The input file cannot be processed due to mismatch in data layout. ***
                
</xsl:text>
			<xsl:value-of select="concat (substring (`$constHyp140, 1, `$blockWidth), '&#xA;')"/>
		</xsl:message>
	</xsl:template>
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']">
		<xsl:variable name="tmp" select="xls:Row[position() > 1][not(xls:Cell[7]/xls:Data = `$master/dty:references/dty:entry/com:countries/com:element)][generate-id (.) = generate-id (key('ccx', xls:Cell[7]/xls:Data)[1])]" />
		<xsl:for-each select="`$tmp">
			<xsl:message>
				<xsl:text>Overshoot: </xsl:text><xsl:text>	</xsl:text><xsl:value-of select="xls:Cell[7]/xls:Data" />
			</xsl:message>
		</xsl:for-each>
		<root>
			<options default="1">
				<prompt>For which country do you want to consolidate the quantities?</prompt>
				<xsl:for-each select="xls:Row[position () > 1][generate-id (.) = generate-id (key('ccx', xls:Cell[7]/xls:Data)[1])][xls:Cell[7]/xls:Data = `$master/dty:references/dty:entry/com:countries/com:element]/xls:Cell[7]/xls:Data">
					<xsl:sort select="." />
					<option id="{position()}">
						<key><xsl:value-of select="." /></key>
						<label><xsl:value-of select="." /></label>
					</option>
				</xsl:for-each>
			</options>
		</root>
	</xsl:template>
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Table" />
	</xsl:template>
	<xsl:template match="/xls:Workbook[xls:Worksheet/@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']/xls:Table" />
	</xsl:template>
	<xsl:template match="/Objects | Objects">
		<xsl:variable name="tmp" select="Object [not(Property [@Name='Empfangsland'] = `$master/dty:references/dty:entry/com:countries/com:element)][generate-id (.) = generate-id (key('ccc', Property [@Name='Empfangsland'])[1])]" />
		<xsl:for-each select="`$tmp">
			<xsl:message>
				<xsl:text>Overshoot: </xsl:text><xsl:text>	</xsl:text><xsl:value-of select="Property [@Name='Empfangsland']" />
			</xsl:message>
		</xsl:for-each>
				<root>
			<options default="1">
				<prompt>For which country do you want to consolidate the quantities?</prompt>
				<xsl:for-each select="Object [generate-id (.) = generate-id (key('ccc', Property [@Name='Empfangsland'])[1])][Property [@Name='Empfangsland'] = `$master/dty:references/dty:entry/com:countries/com:element]/Property [@Name='Empfangsland']">
					<xsl:sort select="." />
					<option id="{position()}">
						<key><xsl:value-of select="." /></key>
						<label><xsl:value-of select="." /></label>
					</option>
				</xsl:for-each>
			</options>
		</root>
	</xsl:template>
	<xsl:template match="/Wrapped">
		<xsl:apply-templates select="Objects" />
	</xsl:template>
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/SAP-DR-Product-Duties.xslt (2022-10-01)
#
# -----------------------------------------------------------------------------------------------
#
[Xml] $script:pdmxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties" 
	exclude-result-prefixes="atom data meta georss gml">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Duties.Batch" select="'VOID'"/>
	<xsl:param name="SAP-DR-Product-Duties.DutyLoadFile" />
	<xsl:param name="SAP-DR-Product-Duties.Product-ContentTypeID" select="'0x01003FAF714C6769BF4FA1B36DCF47ED659702'" />
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry"/>
	<xsl:key name="dref" use="." match="/atom:feed/atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element"/>
	<xsl:key name="duty" use="data:Id" match="/atom:feed/atom:entry/atom:content/meta:properties"/>
	<xsl:template match="data:*" mode="final">
		<xsl:element name="{local-name()}">
			<xsl:apply-templates />
		</xsl:element>
	</xsl:template>
	<xsl:template match="data:REACH | data:Description1 | data:Material | data:ItemListId | data:Duty_x002d_ListId | data:Part_x002d_ListId | data:FileSystemObjectType | data:Id | data:ServerRedirectedEmbedUri | data:ServerRedirectedEmbedUrl | data:ID | data:ContentTypeId | data:Title | data:Modified | data:Created | data:AuthorId | data:EditorId | data:OData__UIVersionString | data:Attachments | data:GUID | data:ComplianceAssetId" mode="final" />
	<xsl:template match="data:element" mode="duty">
		<xsl:param name="entry"/>
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:variable name="tmp" select="." />
		<xsl:if test="`$duties and . = `$duties">
			<xsl:if test="`$verbose">
				<xsl:comment>
					<xsl:text>
				Recycling Duty:
				===========
					
					Origin (material where it is directly attached to):
						- Material:	</xsl:text><xsl:value-of select="`$entry/atom:content/meta:properties/data:Material" /><xsl:text>
						- Description:	</xsl:text><xsl:value-of select="`$entry/atom:content/meta:properties/data:Description1" /><xsl:if test="`$entry/atom:content/meta:properties/data:Pieces [not (@meta:null)]"><xsl:text>
						- Lot size:		</xsl:text><xsl:value-of select="`$entry/atom:content/meta:properties/data:Pieces" /><xsl:text> (order quantity will be diveded by lot size)</xsl:text></xsl:if><xsl:text>
						- Created:		</xsl:text><xsl:value-of select="`$entry/atom:content/meta:properties/data:Modified" /><xsl:text>
						
					Details:
						- </xsl:text><xsl:value-of select="`$duties[. = current()]/parent::meta:properties/data:Duty" /><xsl:text>
						- </xsl:text><xsl:value-of select="`$duties[. = current()]/parent::meta:properties/data:Description12" /><xsl:text>
						- </xsl:text><xsl:value-of select="`$duties[. = current()]/parent::meta:properties/data:Batch" /><xsl:text>
						
					Chain of recursively inspected items (Sharepoint IDs):
						- </xsl:text><xsl:value-of select="`$loop" /><xsl:text>
						
		</xsl:text>
				</xsl:comment>
			</xsl:if>
			<prd:duty SPKey="{.}">
				<prd:data>
					<xsl:variable name="base" select="`$entry/atom:content/meta:properties" />
					<xsl:apply-templates select="`$base/child::node() [not(@meta:null = 'true')]" mode="final" />
				</prd:data>
			</prd:duty>
		</xsl:if>
	</xsl:template>
	<xsl:template match="data:element" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="key('entry',.)" mode="dig">
			<xsl:with-param name="duties" select="`$duties" />
			<xsl:with-param name="loop" select="`$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:element">
		<com:element>
			<xsl:value-of select="."/>
		</com:element>
	</xsl:template>
	<xsl:template match="data:Duty_x002d_ListId">
		<xsl:param name="entry"/>
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="data:element" mode="duty">
			<xsl:with-param name="entry" select="`$entry"/>
			<xsl:with-param name="duties" select="`$duties" />
			<xsl:with-param name="loop" select="`$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="data:element" mode="dig">
			<xsl:with-param name="duties" select="`$duties" />
			<xsl:with-param name="loop" select="`$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:ItemListId">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<prd:duties>
			<xsl:apply-templates select="data:element" mode="dig">
				<xsl:with-param name="duties" select="`$duties" />
				<xsl:with-param name="loop" select="`$loop" />
			</xsl:apply-templates>
		</prd:duties>
	</xsl:template>
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:param name="count" />
		<xsl:if test="`$count > 1">
			<xsl:apply-templates select=".">
				<xsl:with-param name="duties" select="`$duties" />
				<xsl:with-param name="loop" select="`$loop" />
				<xsl:with-param name="count" select="`$count - 1" />
			</xsl:apply-templates>
		</xsl:if>
		<xsl:apply-templates select="key('entry',.)" mode="dig">
			<xsl:with-param name="duties" select="`$duties" />
			<xsl:with-param name="loop" select="`$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Country">
		<com:countries>
			<xsl:apply-templates select="data:element"/>
		</com:countries>
	</xsl:template>
	<xsl:template match="atom:entry" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:choose>
			<xsl:when test="contains(`$loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message>
					<xsl:text>[FATAL] SAP-DR-Product-Duties.xslt (line 163): Loop detected in chain </xsl:text><xsl:value-of select="`$loop" /><xsl:text> Skipping.</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId [not(@meta:null = 'true')]">
					<xsl:with-param name="duties" select="`$duties" />
					<xsl:with-param name="loop" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
					<xsl:with-param name="count" select="atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:ItemListId" mode="dig">
					<xsl:with-param name="duties" select="`$duties" />
					<xsl:with-param name="loop" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId" mode="dig">
					<xsl:with-param name="duties" select="`$duties" />
					<xsl:with-param name="loop" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Duty_x002d_ListId">
					<xsl:with-param name="duties" select="`$duties" />
					<xsl:with-param name="loop" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
					<xsl:with-param name="entry" select="."/>
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="atom:content/meta:properties/data:Material">
		<prd:material>
			<xsl:value-of select="." />
		</prd:material>
	</xsl:template>
	<xsl:template match="atom:content/meta:properties/data:Pieces [not (@meta:null)]">
		<prd:lotsize>
			<xsl:value-of select="." />
		</prd:lotsize>
	</xsl:template>
	<xsl:template match="atom:entry">
		<xsl:param name="duties" />
		<xsl:if test="`$verbose">
			<xsl:comment>
				<xsl:text>
		##########################################################################################
		##########################################################################################
										
					Catalogue product:
					==============
						- Material:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Material" /><xsl:text>
						- Description:	</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Description1" /><xsl:if test="atom:content/meta:properties/data:Pieces [not (@meta:null)]"><xsl:text>
						- Lot size:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Pieces" /><xsl:text> (order quantity will be diveded by lot size)</xsl:text></xsl:if><xsl:text>
						- Created:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Modified" /><xsl:text>
						
	</xsl:text>
			</xsl:comment>
		</xsl:if>
		<prd:product SPKey="{atom:content/meta:properties/data:Id}">
			<xsl:apply-templates select="atom:content/meta:properties/data:Material" />
			<xsl:apply-templates select="atom:content/meta:properties/data:Pieces" />
			<xsl:apply-templates select="atom:content/meta:properties/data:ItemListId">
				<xsl:with-param name="duties" select="`$duties" />
				<xsl:with-param name="loop" select="concat('[', atom:content/meta:properties/data:Id, ']')" />
			</xsl:apply-templates>
		</prd:product>
	</xsl:template>
	<xsl:template match="atom:entry" mode="init">
		<dty:entry SPKey="{atom:content/meta:properties/data:Id}">
			<dty:code>
				<xsl:value-of select="atom:content/meta:properties/data:Duty"/>
			</dty:code>
			<dty:batch>
				<xsl:value-of select="atom:content/meta:properties/data:Batch"/>
			</dty:batch>
			<dty:label>
				<xsl:value-of select="atom:content/meta:properties/data:Description12"/>
			</dty:label>
			<dty:rank>
				<xsl:value-of select="atom:content/meta:properties/data:Rank"/>
			</dty:rank>
			<xsl:apply-templates select="atom:content/meta:properties/data:Country"/>
		</dty:entry>
	</xsl:template>
	<xsl:template match="/atom:feed">
		<xsl:variable name="drefs" select="atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element"/>
		<xsl:variable name="used" select="document(`$SAP-DR-Product-Duties.DutyLoadFile)/atom:feed/atom:entry[atom:content/meta:properties/data:Id = `$drefs][atom:content/meta:properties/data:Batch = `$SAP-DR-Product-Duties.Batch]"/>
		<xsl:comment>
			<xsl:text>
	++++++++++++++++++++++++++++++++++++++++++++++++++++
	
		Debug: </xsl:text><xsl:value-of select="`$debug" />
			<xsl:text>
			
		Verbose: </xsl:text><xsl:value-of select="`$verbose" />
			<xsl:text>
			
	++++++++++++++++++++++++++++++++++++++++++++++++++++		
</xsl:text>
		</xsl:comment>
		<prd:root batch="{`$SAP-DR-Product-Duties.Batch}">
			<dty:references>
				<xsl:attribute name="SPKeys">
					<xsl:for-each select="`$used/atom:content/meta:properties/data:Id">
						<xsl:value-of select="."/>
							<xsl:if test="position () != last ()">
								<xsl:text>,</xsl:text>
							</xsl:if>
						</xsl:for-each>
					</xsl:attribute>
				<xsl:apply-templates select="`$used" mode="init"/>
			</dty:references>
			<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, `$SAP-DR-Product-Duties.Product-ContentTypeID)]">
				<xsl:with-param name="duties" select="`$used/atom:content/meta:properties/data:Id" />
			</xsl:apply-templates>
		</prd:root>
	</xsl:template>
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/SAP-DR-Recycling-Preprocess.xslt (2023-01-06)
#
#
[Xml] $script:intxsl =@"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:ns1="http://www.w3.org/2005/Atom" 
	xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml" 
	xmlns:xls="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:v="urn:schemas-microsoft-com:vml" 
	xmlns:x="urn:schemas-microsoft-com:office:excel" 
	xmlns:xlink="http://www.w3.org/1999/xlink" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:msg="http://www.rothenberger.com/productcompliance/recycling/messages" 
	exclude-result-prefixes="com d georss gml m msg ns1 o prd ss v x xls xlink">
	<xsl:param name="debug" select="0" />
	<xsl:param name="verbose" select="0" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.MasterData" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.Country" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.NumberGroupSeparator" select="'.'" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.NumberDecimalSeparator" select="','" />
	<xsl:variable name="master" select="document (`$SAP-DR-Recycling-Preprocess.MasterData)/prd:root" />
	<xsl:variable name="product" select="`$master/prd:product" />
	<xsl:variable name="constBlk140" select="'                                                                                                                                            '" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="constDot140" select="'............................................................................................................................................'" />
	<xsl:variable name="blockWidth" select="95" />
	<xsl:variable name="indent1" select="10" />
	<xsl:variable name="indent2" select="22" />
	<xsl:variable name="indent3" select="45" />
	<xsl:variable name="indent4" select="70" />
	<xsl:key name="mat-xml" use="concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data)" match="xls:Row" />
	<xsl:key name="mat-csv" use="concat (Property[@Name='Material'],'-', Property[@Name='Empfangsland'])" match="Object" />
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<xsl:output method="xml" encoding="UTF-8" indent="yes" standalone="yes"  version="1.0" />
	<xsl:template match="/*">
		<xsl:message>
			<xsl:value-of select="concat (substring (`$constHyp140, 1, `$blockWidth), '&#xA;')"/>
			<xsl:text>
 *** ERROR: The input file cannot be processed due to mismatch in data layout. ***
                
</xsl:text>
			<xsl:value-of select="concat (substring (`$constHyp140, 1, `$blockWidth), '&#xA;')"/>
		</xsl:message>
	</xsl:template>
	<xsl:template name="rec-count-lot-mismatch-xml">
		<xsl:param name="rows" />
		<xsl:param name="hits" select="0" />
		<xsl:choose>
			<xsl:when test="`$rows [1]">
				<xsl:choose>
					<xsl:when test="`$product [prd:material = `$rows[1]/xls:Cell[1]/xls:Data]/prd:lotsize and not (0 = `$rows[1]/xls:Cell[5]/xls:Data mod `$product [prd:material = `$rows[1]/xls:Cell[1]/xls:Data]/prd:lotsize)">
						<xsl:call-template name="rec-count-lot-mismatch-xml">
							<xsl:with-param name="rows" select="`$rows [position () > 1]" />
							<xsl:with-param name="hits" select="`$hits + 1" />
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="rec-count-lot-mismatch-xml">
							<xsl:with-param name="rows" select="`$rows [position () > 1]" />
							<xsl:with-param name="hits" select="`$hits" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="`$hits" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rec-count-lot-mismatch-csv">
		<xsl:param name="objects" />
		<xsl:param name="hits" select="0" />
		<xsl:choose>
			<xsl:when test="`$objects [1]">
				<xsl:variable name="material" select="`$objects[1]/Property[@Name='Material']" />
				<xsl:choose>
					<xsl:when test="`$product [prd:material = `$material]/prd:lotsize and not (0 = `$objects[1]/Property[@Name='Fakturierte Menge'] mod `$product [prd:material = `$material]/prd:lotsize)">
						<xsl:call-template name="rec-count-lot-mismatch-csv">
							<xsl:with-param name="objects" select="`$objects [position () > 1]" />
							<xsl:with-param name="hits" select="`$hits + 1" />
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="rec-count-lot-mismatch-csv">
							<xsl:with-param name="objects" select="`$objects [position () > 1]" />
							<xsl:with-param name="hits" select="`$hits" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="`$hits" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="xls:Row">
		<item xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<material>
				<xsl:value-of select="xls:Cell[1]/xls:Data" />
			</material>
			<kurztext>
				<xsl:value-of select="xls:Cell[2]/xls:Data" />
			</kurztext>
			<country>
				<xsl:value-of select="xls:Cell[7]/xls:Data" />
			</country>
			<units>
				<xsl:choose>
					<xsl:when test="`$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize > 1">
						<xsl:attribute name="lot">
							<xsl:value-of select="`$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize" />
						</xsl:attribute>
						<xsl:value-of select="ceiling (sum (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))/xls:Cell[5]/xls:Data) div `$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="sum (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))/xls:Cell[5]/xls:Data)" />
					</xsl:otherwise>
				</xsl:choose>
			</units>
		</item>
	</xsl:template>
	<xsl:template match="xls:Table [xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']">
		<report xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<xsl:if test="not (`$product)">
				<xsl:message terminate="yes"> 
					<xsl:text>FATAL: Product master data missing. Aborting.</xsl:text>
				</xsl:message>
			</xsl:if>
			<xsl:variable name="hits">
				<xsl:call-template name="rec-count-lot-mismatch-xml">
					<xsl:with-param name="rows" select="xls:Row [position () > 1][xls:Cell[7]/xls:Data = `$SAP-DR-Recycling-Preprocess.Country]" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="`$hits > 0">
				<messages xmlns="http://www.rothenberger.com/productcompliance/recycling/messages" number="{`$hits}">
					<caption>
						<xsl:text>Lines in import file where 'quantity' does not match the lot size:</xsl:text>
					</caption>
					<xsl:for-each select="xls:Row [position () > 1]">
						<xsl:if test="`$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize and not (0 = xls:Cell[5]/xls:Data mod `$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)">
							<message>
								<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1 - 1 - string-length (format-number (position(), '#.##0'))), format-number (position(), '#.##0'), ' ', xls:Cell [1]/xls:Data, substring (`$constBlk140, 1, `$indent2 - `$indent1 - string-length (xls:Cell [1]/xls:Datal)), ' ', xls:Cell [2]/xls:Data, ' pcs.: ', xls:Cell [5]/xls:Data, ' lot size: ', `$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)" />
							</message>
						</xsl:if>
					</xsl:for-each>
				</messages>
			</xsl:if>
			<xsl:apply-templates select="xls:Row [position () > 1][xls:Cell[7]/xls:Data = `$SAP-DR-Recycling-Preprocess.Country][generate-id (.) = generate-id (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))[1])]" />
		</report>
	</xsl:template>
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Table" />
	</xsl:template>
	<xsl:template match="/xls:Workbook[xls:Worksheet/@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']" />
	</xsl:template>
	<xsl:template match="Object [Property/@Name='Material' and Property/@Name='Fakturierte Menge' and Property/@Name='Empfangsland']" mode="wrapped">
		<xsl:variable name="material" select="Property[@Name='Material']" />
		<item xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<material>
				<xsl:value-of select="`$material" />
			</material>
			<kurztext>
				<xsl:value-of select="Property[@Name='Materialkurztext']" />
			</kurztext>
			<country>
				<xsl:value-of select="Property[@Name='Empfangsland']" />
			</country>
			<units>
				<xsl:choose>
					<xsl:when test="`$product [prd:material = `$material]/prd:lotsize > 1">
						<xsl:attribute name="lot">
							<xsl:value-of select="`$product [prd:material = `$material]/prd:lotsize" />
						</xsl:attribute>
						<xsl:value-of select="ceiling (sum (key('mat-csv', concat (Property[@Name='Material'], '-', Property[@Name='Empfangsland']))/Property[@Name='Fakturierte Menge']) div `$product [prd:material = `$material]/prd:lotsize)" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="sum (key('mat-csv', concat (Property[@Name='Material'], '-', Property[@Name='Empfangsland']))/Property[@Name='Fakturierte Menge'])" />
					</xsl:otherwise>
				</xsl:choose>
			</units>
		</item>
	</xsl:template>
	<xsl:template match="Objects" mode="wrapped">
		<report xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<xsl:if test="not (`$product)">
				<xsl:message terminate="yes"> 
					<xsl:text>FATAL: Product master data missing. Aborting.</xsl:text>
				</xsl:message>
			</xsl:if>
			<xsl:variable name="hits">
				<xsl:call-template name="rec-count-lot-mismatch-csv">
					<xsl:with-param name="objects" select="Object[Property[@Name='Empfangsland'] = `$SAP-DR-Recycling-Preprocess.Country]" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="`$hits > 0">
				<messages xmlns="http://www.rothenberger.com/productcompliance/recycling/messages" number="{`$hits}">
					<caption>
						<xsl:text>Lines in import file where 'quantity' does not match the lot size:</xsl:text>
					</caption>
					<xsl:for-each select="Object [Property [@Name='Empfangsland'] = `$SAP-DR-Recycling-Preprocess.Country]">
						<xsl:variable name="material" select="Property [@Name='Material']" />
						<xsl:variable name="kurztext" select="Property [@Name='Materialkurztext']" />
						<xsl:variable name="amount" select="Property [@Name='Fakturierte Menge']" />
						<xsl:if test="`$product [prd:material = `$material]/prd:lotsize and not (0 = `$amount mod `$product [prd:material = `$material]/prd:lotsize)">
							<message>
								<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1 - 1 - string-length (format-number (position(), '#.##0'))), format-number (position(), '#.##0'), ' ', `$material, substring (`$constBlk140, 1, `$indent2 - `$indent1 - string-length (`$material)), ' ', `$kurztext, ' pcs.: ', `$amount, ' lot size: ', `$product [prd:material = `$material]/prd:lotsize)" />
							</message>
						</xsl:if>
					</xsl:for-each>
				</messages>
			</xsl:if>
			<xsl:apply-templates select="Object[generate-id (.) = generate-id (key('mat-csv', concat (Property[@Name='Material'], '-', `$SAP-DR-Recycling-Preprocess.Country))[1])]" mode="wrapped">
				<xsl:sort select="Property[@Name='Material']" />
			</xsl:apply-templates>
		</report>
	</xsl:template>
	<xsl:template match="/Wrapped">
		<xsl:apply-templates select="Objects" mode="wrapped" />
	</xsl:template>
	<xsl:template match="Property" />
	<xsl:template match="Property [@Name='Fakturierte Menge']">
		<xsl:param name="amount" />
		<xsl:copy>
			<xsl:copy-of select="@Name" />
			<xsl:value-of select="`$amount" />
		</xsl:copy>
	</xsl:template>
	<xsl:template match="Property [@Name='Material'] | Property[@Name='Materialkurztext'] | Property [@Name='Empfangsland']">
		<xsl:copy>
			<xsl:copy-of select="@Name" />
			<xsl:value-of select="normalize-space (.)" />
		</xsl:copy>
	</xsl:template>
	<xsl:template match="Object">
		<xsl:variable name="amount">
			<xsl:choose>
				<xsl:when test="substring-after (Property [@Name='Fakturierte Menge'], `$SAP-DR-Recycling-Preprocess.NumberDecimalSeparator)">
					<xsl:message>
						<xsl:text>WARNING: Sales volume should not be decimal formatted numbers (adjusted)</xsl:text>
					</xsl:message>
					<xsl:value-of select="translate (substring-before (Property [@Name='Fakturierte Menge'], `$SAP-DR-Recycling-Preprocess.NumberDecimalSeparator), concat ('0123456789', `$SAP-DR-Recycling-Preprocess.NumberGroupSeparator), '0123456789')" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="translate (Property [@Name='Fakturierte Menge'], concat ('0123456789', `$SAP-DR-Recycling-Preprocess.NumberGroupSeparator), '0123456789')" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:if test="not (number (`$amount) = 0)">
			<xsl:copy>
				<xsl:apply-templates select="Property">
					<xsl:with-param name="amount" select="`$amount" />
				</xsl:apply-templates>
			</xsl:copy>
		</xsl:if>
	</xsl:template>
	<xsl:template match="/Objects">
		<Wrapped>
			<xsl:copy>
				<xsl:apply-templates select="Object" />
			</xsl:copy>
		</Wrapped>
	</xsl:template>
</xsl:stylesheet>
"@
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/SAP-DR-Recycling-Calculate.xslt (2022-10-01)
#
[Xml] $script:conxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:ns1="http://www.w3.org/2005/Atom" 
	xmlns:xlink="http://www.w3.org/1999/xlink" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:msg="http://www.rothenberger.com/productcompliance/recycling/messages" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty prd">
	<xsl:param name="debug" select="0" />
	<xsl:param name="verbose" select="0" />
	<xsl:param name="SAP-DR-Recycling-Converter.Country" />
	<xsl:param name="SAP-DR-Recycling-Converter.MasterData" />
	<xsl:variable name="indent1" select="10" />
	<xsl:variable name="indent2" select="22" />
	<xsl:variable name="indent3" select="45" />
	<xsl:variable name="indent4" select="70" />
	<xsl:variable name="constBlk140" select="'                                                                                                                                            '" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="constDot140" select="'............................................................................................................................................'" />
	<xsl:variable name="blockWidth" select="95" />
	<xsl:variable name="master" select="document (`$SAP-DR-Recycling-Converter.MasterData)/prd:root" />
	<xsl:variable name="batch" select="`$master/@batch" />
	<xsl:variable name="filter" select="`$master/dty:references/dty:entry [com:countries/com:element = `$SAP-DR-Recycling-Converter.Country]/@SPKey" />
	<xsl:variable name="duties" select="`$master/prd:product/prd:duties/prd:duty[@SPKey = `$filter]" />
	<xsl:variable name="products" select="`$master/prd:product" />
	<xsl:key name="mat" use="com:material" match="com:item" />
	<xsl:key name="cat" use="@SPKey" match="prd:product/prd:duties/prd:duty" />
	<xsl:key name="prd" use="prd:product/prd:duties/prd:duty/@SPKey" match="prd:product" />
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<xsl:output method="text" encoding="UTF-8" />
	<xsl:template name="format">
		<xsl:param name="hdr" />
		<xsl:param name="lab" />
		<xsl:param name="pcs" />
		<xsl:param name="wgt" />
		<xsl:variable name="label" select="substring(`$lab, 1, 20)" />
		<xsl:variable name="pieces" select="format-number(`$pcs, '#.##0')" />
		<xsl:variable name="weight" select="format-number(`$wgt, '#.##0,000')" />
		<xsl:if test="`$hdr">
			<xsl:value-of select="concat(substring (`$constBlk140, 1, `$indent1), `$hdr, ':', '&#xA;&#xA;')" />
		</xsl:if>
		<xsl:value-of select="concat(substring(`$constBlk140, 1, `$indent3), `$lab, ' ', substring (`$constDot140, 1, `$indent4 - `$indent3 - string-length (`$lab) - string-length (`$pieces)), ' ', `$pieces, ' pcs.', substring (`$constBlk140, 1, 12 - string-length (`$weight)), `$weight, ' kg', '&#xA;&#xA;')" />
	</xsl:template>
	<xsl:template name="batt-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="scr" select="0" />
		<xsl:param name="pcs" select="0" />
		<xsl:param name="ces" select="0" />
		<xsl:variable name="pos" select="`$duties [1]" />
		<xsl:choose>
			<xsl:when test="`$pos">
				<xsl:variable name="rows" select="`$table [com:material = `$pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="batt-rec">
					<xsl:with-param name="header" select="`$header" />
					<xsl:with-param name="duties" select="`$duties [position() > 1]" />
					<xsl:with-param name="table" select="`$table" />
					<xsl:with-param name="inv" select="`$inv" />
					<xsl:with-param name="cnt" select="`$cnt + sum(`$rows/com:units)" />
					<xsl:with-param name="scr" select="`$scr + sum(`$rows/com:units) * `$pos/prd:data/Weight" />
					<xsl:with-param name="pcs" select="`$pcs + sum(`$rows/com:units)" />	
					<!-- xsl:with-param name="pcs" select="`$pcs + sum(`$rows/com:units) * `$pos/prd:data/Pieces" /-->
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="`$verbose or `$cnt > 0 or `$scr > 0">
					<xsl:call-template name="format">
						<xsl:with-param name="hdr" select="`$header" />
						<xsl:with-param name="lab" select="'Total:'" />
						<xsl:with-param name="pcs" select="`$pcs" />
						<xsl:with-param name="wgt" select="`$scr" />
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="weee-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="scr" select="0" />
		<xsl:variable name="pos" select="`$duties [1]" />
		<xsl:choose>
			<xsl:when test="`$pos">
				<xsl:variable name="rows" select="`$table [com:material = `$pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="weee-rec">
					<xsl:with-param name="header" select="`$header" />
					<xsl:with-param name="duties" select="`$duties [position() > 1]" />
					<xsl:with-param name="table" select="`$table" />
					<xsl:with-param name="inv" select="`$inv" />
					<xsl:with-param name="cnt" select="`$cnt + sum(`$rows/com:units)" />
					<xsl:with-param name="scr" select="`$scr + sum(`$rows/com:units) * `$pos/prd:data/Weight" />	
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="`$verbose or `$cnt > 0 or `$scr > 0">
					<xsl:call-template name="format">
						<xsl:with-param name="hdr" select="`$header" />
						<xsl:with-param name="lab" select="'Total:'" />
						<xsl:with-param name="pcs" select="`$cnt" />
						<xsl:with-param name="wgt" select="`$scr" />
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="tvvv-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="alu" select="0" />
		<xsl:param name="alu-cnt" select="0" />
		<xsl:param name="iron" select="0" />
		<xsl:param name="iron-cnt" select="0" />
		<xsl:param name="ppk" select="0" />
		<xsl:param name="ppk-cnt" select="0" />
		<xsl:param name="plast" select="0" />
		<xsl:param name="plast-cnt" select="0" />
		<xsl:param name="tin" select="0" />
		<xsl:param name="tin-cnt" select="0" />
		<xsl:variable name="pos" select="`$duties [1]" />
		<xsl:choose>
			<xsl:when test="`$pos">
				<xsl:variable name="rows" select="`$table [com:material = `$pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="tvvv-rec">
					<xsl:with-param name="header" select="`$header" />
					<xsl:with-param name="duties" select="`$duties [position() > 1]" />
					<xsl:with-param name="table" select="`$table" />
					<xsl:with-param name="inv" select="`$inv" />
					<xsl:with-param name="cnt" select="`$cnt + sum(`$rows/com:units)" />
					<xsl:with-param name="alu" select="`$alu + sum(`$rows/com:units) * `$pos/prd:data/VV_x002d_Alu" />
					<xsl:with-param name="iron" select="`$iron + sum(`$rows/com:units) * `$pos/prd:data/VV_x002d_Steel" />
					<xsl:with-param name="ppk" select="`$ppk + sum(`$rows/com:units) * `$pos/prd:data/Paper" />
					<xsl:with-param name="plast" select="`$plast + sum(`$rows/com:units) * `$pos/prd:data/VV_x002d_Plastic" />
					<xsl:with-param name="tin" select="`$tin + sum(`$rows/com:units) * `$pos/prd:data/VV_x002d_Tinplate" />
					<xsl:with-param name="alu-cnt">
						<xsl:choose>
							<xsl:when test="`$pos/prd:data/VV_x002d_Alu > 0">
								<xsl:value-of select="`$alu-cnt + sum(`$rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$alu-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="iron-cnt">
						<xsl:choose>
							<xsl:when test="`$pos/prd:data/VV_x002d_Steel > 0">
								<xsl:value-of select="`$iron-cnt + sum(`$rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$iron-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="ppk-cnt">
						<xsl:choose>
							<xsl:when test="`$pos/prd:data/Paper > 0">
								<xsl:value-of select="`$ppk-cnt + sum(`$rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$ppk-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="plast-cnt">
						<xsl:choose>
							<xsl:when test="`$pos/prd:data/VV_x002d_Plastic > 0">
								<xsl:value-of select="`$plast-cnt + sum(`$rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$plast-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="tin-cnt">
						<xsl:choose>
							<xsl:when test="`$pos/prd:data/VV_x002d_Tinplate > 0">
								<xsl:value-of select="`$tin-cnt + sum(`$rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$tin-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="format">
					<xsl:with-param name="hdr" select="`$header" />
					<xsl:with-param name="lab" select="'Paper'" />
					<xsl:with-param name="wgt" select="`$ppk * 0.001" />
					<xsl:with-param name="pcs" select="`$ppk-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Tinplate'" />
					<xsl:with-param name="wgt" select="`$tin * 0.001" />
					<xsl:with-param name="pcs" select="`$tin-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Alu'" />
					<xsl:with-param name="wgt" select="`$alu * 0.001" />
					<xsl:with-param name="pcs" select="`$alu-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Iron'" />
					<xsl:with-param name="wgt" select="`$iron * 0.001" />
					<xsl:with-param name="pcs" select="`$iron-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Plastics'" />
					<xsl:with-param name="wgt" select="`$plast* 0.001" />
					<xsl:with-param name="pcs" select="`$plast-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Total'" />
					<xsl:with-param name="wgt" select="(`$alu + `$iron + `$tin + `$ppk + `$plast) * 0.001" />
					<xsl:with-param name="pcs" select="`$cnt" />
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="com:item" mode="check">
		<xsl:variable name="total" select="sum (key('mat', com:material)[com:country = `$SAP-DR-Recycling-Converter.Country]/com:units)"/>
		<xsl:variable name="col1w" select="`$indent1"/>
		<xsl:variable name="col2w" select="`$indent2 - `$indent1"/>
		<xsl:variable name="col3w" select="40"/>
		<xsl:variable name="col4w" select="8"/>
		<xsl:variable name="col5w" select="6"/>
		<xsl:variable name="col6w" select="8"/>
		<xsl:value-of select="concat (substring (`$constBlk140, 1, `$col1w), com:material, substring (`$constBlk140, 1, `$col2w - string-length (com:material)), substring (com:kurztext, 1, `$col3w))"/>
		<xsl:value-of select="concat (substring (`$constBlk140, 1, `$col5w + `$col4w + `$col3w - string-length (substring (com:kurztext, 1, `$col3w))), substring (`$constBlk140, 1, `$col6w - string-length (format-number (`$total, '#.##0'))), format-number (`$total, '#.##0'), ' pcs.&#xA;')" />
	</xsl:template>
	<xsl:template match="msg:messages">
		<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1), '********* ', msg:caption,'  *********', '&#xA;&#xA;')"/>
		<xsl:for-each select="child::msg:message">
			<xsl:value-of select="concat(. ,'&#xA;')" />
		</xsl:for-each>
		<xsl:value-of select="concat ('&#xA;', substring(`$constHyp140, 1, `$blockWidth), '&#xA;&#xA;')" />
	</xsl:template>
	<xsl:template match="com:report" mode="check">
		<xsl:variable name="nom" select="com:item [com:country = `$SAP-DR-Recycling-Converter.Country][not(com:material = `$duties/parent::prd:duties/parent::prd:product/prd:material)][generate-id (key('mat', com:material)[com:country = `$SAP-DR-Recycling-Converter.Country][1]) = generate-id (.)]" />
		<xsl:if test="`$nom">
			<xsl:if test="`$nom [not(com:material = `$master/prd:product/prd:material)]">
				<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1), '********* Catalogue numbers in import file with no master data *********', '&#xA;&#xA;')"/>
				<xsl:apply-templates select="`$nom [not(com:material = `$master/prd:product/prd:material)]" mode="check">
					<xsl:sort select="com:material" />
				</xsl:apply-templates>
				<xsl:value-of select="'&#xA;'" />
			</xsl:if>
			<xsl:if test="`$verbose">
				<xsl:if test="`$nom [com:material = `$master/prd:product/prd:material]">
					<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1), '********* Catalogue numbers in import file with no duties in current batch ', `$batch, ' and country ', `$SAP-DR-Recycling-Converter.Country ,'  *********', '&#xA;&#xA;')"/>
					<xsl:apply-templates select="`$nom [com:material = `$master/prd:product/prd:material]" mode="check">
						<xsl:sort select="com:material" />
					</xsl:apply-templates>
					<xsl:value-of select="'&#xA;'" />
				</xsl:if>
			</xsl:if>
			<xsl:if test="not(msg:messages)">
				<xsl:value-of select="concat (substring(`$constHyp140, 1, `$blockWidth), '&#xA;&#xA;')" />
			</xsl:if>
		</xsl:if>
		<xsl:apply-templates select="msg:messages" />
	</xsl:template>
	<xsl:template match="com:report" mode="run">
		<xsl:variable name="here" select="current()" />
		<xsl:for-each select="`$duties [generate-id (key('cat', @SPKey)[1]) = generate-id (.)]">
			<xsl:sort select="`$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank" data-type="number" order="ascending" />
			<xsl:variable name="filter" select="key('cat',@SPKey)" />
			<xsl:variable name="rowx" select="`$here/com:item [com:country = `$SAP-DR-Recycling-Converter.Country][com:material = `$filter/parent::prd:duties/parent::prd:product/prd:material]" />
			<xsl:choose>
				<xsl:when test="`$batch = 'BATT'">
					<xsl:call-template name="batt-rec">
						<xsl:with-param name="header" select="concat ('[Line ', `$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank, '] ', `$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label)" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="`$rowx" />
						<xsl:with-param name="inv" select="count(`$rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:when test="`$batch = 'TVVV'">
					<xsl:call-template name="tvvv-rec">
						<xsl:with-param name="header" select="`$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="`$rowx" />
						<xsl:with-param name="inv" select="count(`$rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:when test="`$batch = 'WEEE'">
					<xsl:call-template name="weee-rec">
						<xsl:with-param name="header" select="concat ('[Line ', `$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank, '] ', `$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label)" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="`$rowx" />
						<xsl:with-param name="inv" select="count(`$rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="concat ('&#xA;', substring (`$constHyp140, 1, `$blockWidth), '&#xA;', '&#xA;')"/>
					<xsl:value-of select="concat (' *** FEHLER: Der Batch &quot;', `$batch, '&quot; ist nicht vorgesehen. ***', '&#xA;')"/>
					<xsl:value-of select="concat ('&#xA;', substring (`$constHyp140, 1, `$blockWidth), '&#xA;', '&#xA;')"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		
	</xsl:template>
	<xsl:template match="/com:report">
		<xsl:value-of select="concat ('&#xA;', substring(`$constHyp140, 1, `$blockWidth), '&#xA;&#xA;')" />
		<xsl:value-of select="concat (substring (`$constBlk140, 1, `$indent1), `$SAP-DR-Recycling-Converter.Country, '&#xA;&#xA;')" />
		<xsl:apply-templates select="." mode="check" />
		<xsl:apply-templates select="." mode="run" />
		<xsl:value-of select="concat ('&#xA;', substring(`$constHyp140, 1, `$blockWidth), '&#xA;')" />
	</xsl:template>
</xsl:stylesheet>
"@
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/SAP-DR-Product-Lookup.xslt (2022-10-01)
#
[Xml] $script:lupxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties" 
	exclude-result-prefixes="atom data meta georss gml">
	<xsl:output method="text" encoding="UTF-8"/>
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.Mode" select="'forward'" />
	<xsl:param name="SAP-DR-Product-Lookup.UnfoldBundles" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.ShowDuties" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.ProductLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.DutyLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.FieldsLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.ContentTypesLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.Material" />
	<xsl:param name="SAP-DR-Product-Lookup.Duty" />
	<xsl:param name="SAP-DR-Product-Lookup.Product-ContentTypeID" select="'0x01003FAF714C6769BF4FA1B36DCF47ED659702'" />
	<xsl:variable name="SAP-DR-Product-Lookup.Products" select="document(`$SAP-DR-Product-Lookup.ProductLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.Duties" select="document(`$SAP-DR-Product-Lookup.DutyLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.Fields" select="document(`$SAP-DR-Product-Lookup.FieldsLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.ContentTypes" select="document(`$SAP-DR-Product-Lookup.ContentTypesLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.ConstBlank140" select="'                                                                                                                                            '"/>
	<xsl:variable name="SAP-DR-Product-Lookup.ConstHyphen140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'"/>
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry" />
	<xsl:key name="back" use="atom:content/meta:properties/data:ItemListId/data:element | atom:content/meta:properties/data:Part_x002d_ListId/data:element | atom:content/meta:properties/data:Reference_x002d_ProductId" match="/atom:feed/atom:entry" />
	<xsl:key name="dref" use="." match="/atom:feed/atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element" />
	<xsl:key name="duty" use="data:Batch" match="/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:template match="data:*" mode="final">
		<xsl:param name="level" />
		<xsl:variable name="name" select="local-name()" />
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="label" select="`$SAP-DR-Product-Lookup.Fields [data:InternalName = `$name]/data:Title" />
			<xsl:with-param name="value" select="." />
		</xsl:call-template>
	</xsl:template>
	<xsl:template match="data:Description1 | data:Material | data:ItemListId | data:Duty_x002d_ListId | data:Part_x002d_ListId | data:Reference_x002d_ProductId | data:FileSystemObjectType | data:Id | data:ServerRedirectedEmbedUri | data:ServerRedirectedEmbedUrl | data:ID | data:ContentTypeId | data:Title | data:Modified | data:Created | data:AuthorId | data:EditorId | data:OData__UIVersionString | data:Attachments | data:GUID | data:ComplianceAssetId" mode="final" />
	<xsl:template match="data:element">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Duty_x002d_ListId">
		<xsl:param name="level" />
		<xsl:param name="entry"/>
		<xsl:variable name= "tmp" select="`$SAP-DR-Product-Lookup.Duties [data:Id = current()/data:element]" />
		<xsl:if test="`$tmp">
			<xsl:for-each select="`$tmp [generate-id(.) = generate-id (key ('duty', data:Batch)[data:Id = current()/data:element][1])]">
				<xsl:call-template name="line-fill">
					<xsl:with-param name="value" select="data:Batch" />
					<xsl:with-param name="level" select="`$level" />
				</xsl:call-template>
				<xsl:for-each select="`$tmp [data:Batch = current()/data:Batch]">
					<xsl:sort select="data:Duty" data-type="text"  order="ascending" />
					<xsl:call-template name="line-fill">
						<xsl:with-param name="level" select="`$level" />
						<xsl:with-param name="label" select="data:Duty" />
						<xsl:with-param name="value" select="data:Description12" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:for-each>
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="`$level" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="data:element">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:param name="count" />
		<xsl:if test="`$SAP-DR-Product-Lookup.UnfoldBundles and `$count > 1">
			<xsl:apply-templates select=".">
				<xsl:with-param name="level" select="`$level" />
				<xsl:with-param name="loop" select="`$loop" />
				<xsl:with-param name="base" select="`$base" />
				<xsl:with-param name="count" select="`$count - 1" />
			</xsl:apply-templates>
		</xsl:if>
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template name="line-fill">
		<xsl:param name="label" select="''"/>
		<xsl:param name="value" select="''"/>
		<xsl:param name="level" />
		<xsl:variable name="border" select="'##'" />
		<xsl:variable name="margin" select="2" />
		<xsl:variable name="indent" select="8" />
		<xsl:variable name="header" select="12" />
		<xsl:variable name="width" select="100" />
		<xsl:variable name="nlabel" select="substring (`$label, 1, `$header)" />
		<xsl:choose>
			<xsl:when test="string-length (`$label) = 0 and string-length (`$value) = 0">
				<xsl:value-of select="concat (`$border, substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$margin), substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$indent * `$level), '+', substring (`$SAP-DR-Product-Lookup.ConstHyphen140, 1, `$width - (`$indent * `$level)),'+&#xA;')" />
			</xsl:when>
			<xsl:when test="string-length (`$label) = 0">
				<xsl:variable name="nvalue" select="substring (`$value, 1, `$width - `$indent * `$level - 2)" />
				<xsl:value-of select="concat (`$border, substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$margin), substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$indent * `$level), '| ', `$nvalue, substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$width - (`$indent * `$level) - string-length (`$nvalue) - 1), '|&#xA;')" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="nvalue" select="substring (`$value, 1, `$width - `$indent * `$level - `$header - 7)" />
				<xsl:value-of select="concat (`$border, substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$margin), substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$indent * `$level), '|  - ', `$nlabel, ': ', substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$header - string-length (`$nlabel)), `$nvalue, substring (`$SAP-DR-Product-Lookup.ConstBlank140, 1, `$width - (`$indent * `$level) - `$header - string-length (`$nvalue) - 6), '|&#xA;')" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rec-probe-duties">
		<xsl:param name="loop" />
		<xsl:param name="batch" />
		<xsl:param name="duty" />
		<xsl:param name="items" />
		<xsl:choose>
			<xsl:when test="count (`$SAP-DR-Product-Lookup.Duties[data:Batch = `$batch or data:Id = `$duty][data:Id = `$items [1]/atom:content/meta:properties/data:Duty_x002d_ListId/data:element]) > 0">
				<xsl:value-of select="1" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="keys" select="`$items [1]/atom:content/meta:properties/data:Reference_x002d_ProductId | `$items [1]/atom:content/meta:properties/data:ItemListId/data:element | `$items[1]/atom:content/meta:properties/data:Part_x002d_ListId/data:element" />
				<xsl:if test="`$keys">
				<xsl:variable name="tmp">
					<xsl:call-template name="rec-probe-duties">
						<xsl:with-param name="items" select="key('entry', `$keys)" />
						<xsl:with-param name="loop" select="concat(`$loop, '-[', `$items [1]/atom:content/meta:properties/data:Id, ']')" />
						<xsl:with-param name="batch" select="`$batch" />
						<xsl:with-param name="duty" select="`$duty" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="`$tmp > 0">
						<xsl:value-of select="1" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="`$items [position () > 1]">
								<xsl:call-template name="rec-probe-duties">
									<xsl:with-param name="loop"  select="`$loop"/>
									<xsl:with-param name="batch" select="`$batch"/>
									<xsl:with-param name="duty" select="`$duty" />
									<xsl:with-param name="items" select="`$items [position () > 1]" />
								</xsl:call-template>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="0" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="atom:entry" mode="text">
		<xsl:param name="level" />
		<xsl:param name="back" select="false ()" />
		<xsl:variable name="shift" select="10" />
		<xsl:variable name="cotyid" select="atom:content/meta:properties/data:ContentTypeId" />
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
		</xsl:call-template>
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="value">
				<xsl:value-of select="`$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name" />
				<xsl:if test="string-length(`$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Description) > 0">
					<xsl:value-of select="concat (' (', `$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Description,')')" />
				</xsl:if>
			</xsl:with-param>
		</xsl:call-template>
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="label" select="'Kurztext'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Description1" />
		</xsl:call-template>
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="label" select="'Material'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Material" />
		</xsl:call-template>
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="label" select="'Attachments'" />
			<xsl:with-param name="value">
				<xsl:choose>
					<xsl:when test="atom:content/meta:properties/data:Attachments = 'true'">
						<xsl:text>[X]</xsl:text>
					</xsl:when>
					<xsl:otherwise>
						<xsl:text>[_]</xsl:text>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:with-param>
		</xsl:call-template>
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="`$level" />
			<xsl:with-param name="label" select="'Erstellt'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Modified" />
		</xsl:call-template>
		<xsl:apply-templates select="atom:content/meta:properties/data:* [not(@meta:null = 'true')]" mode="final">
			<xsl:with-param name="level" select="`$level" />
		</xsl:apply-templates>
		<xsl:if test="`$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name = 'Product' or `$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name = 'Notified-Product'">
			<xsl:variable name="batt">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'BATT'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="tvvv">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'TVVV'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="weee">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'WEEE'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="scip">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'SCIP'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="`$level" />
				<xsl:with-param name="value" select="concat ('   BATT [', translate(`$batt, '01', '_X'), ']      TVVV [', translate(`$tvvv, '01', '_X'), ']      WEEE [', translate(`$weee, '01', '_X'), ']      SCIP [', translate(`$scip, '01', '_X'),']')" />
			</xsl:call-template>
		</xsl:if>
		<xsl:if test="`$back or not(`$SAP-DR-Product-Lookup.ShowDuties and atom:content/meta:properties/data:Duty_x002d_ListId/data:element)">
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="`$level" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template match="atom:entry">
		<xsl:param name="level" />
		<xsl:param name="indent" />
		<xsl:param name="loop" select="'>>'" />
		<xsl:param name="base" select="." />
		<xsl:param name="trace" select="''" />
		<xsl:variable name="next" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
		<xsl:choose>
			<xsl:when test="contains(`$loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 386): Loop detected in chain </xsl:text><xsl:value-of select="`$next" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'attachments'">
						<xsl:variable name="props" select="atom:content/meta:properties" />
						<xsl:if test="not (contains(`$loop, concat ('[', atom:content/meta:properties/data:Id, ']'))) and atom:content/meta:properties/data:Attachments = 'true'">
							<xsl:value-of select="concat (`$base/atom:content/meta:properties/data:Id, ';', translate (`$base/atom:content/meta:properties/data:Material, ';',','), ';', translate (`$base/atom:content/meta:properties/data:Description1, ';', ','), ';', `$props/data:Id, ';', translate (`$props/data:Material, ';[](){}|', ','), ';', `$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = `$props/data:ContentTypeId]/data:Name, ';', translate (`$props/data:Description1, ';', ','), '&#xA;')" />
						</xsl:if>
					</xsl:when>
					<xsl:otherwise>
						<xsl:apply-templates select="." mode="text">
							<xsl:with-param name="level" select="`$level" />
						</xsl:apply-templates>
						<xsl:if test="`$SAP-DR-Product-Lookup.ShowDuties">
							<xsl:apply-templates select="atom:content/meta:properties/data:Duty_x002d_ListId">
								<xsl:with-param name="level" select="`$level" />
								<xsl:with-param name="entry" select="."/>
							</xsl:apply-templates>
						</xsl:if>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId">
					<xsl:with-param name="level" select="`$level + 1" />
					<xsl:with-param name="loop" select="`$next" />
					<xsl:with-param name="base" select="`$base" />
					<xsl:with-param name="count" select="atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId | atom:content/meta:properties/data:ItemListId">
					<xsl:with-param name="level" select="`$level + 1" />
					<xsl:with-param name="loop" select="`$next" />
					<xsl:with-param name="base" select="`$base" />
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="track-down">
		<xsl:param name="level" select="0" />
		<xsl:param name="chain" />
		<xsl:variable name="key" select="substring-before (substring-after (`$chain, '['), ']')" />
		<xsl:if test="not (contains(`$key, 'x'))">
			<xsl:apply-templates select="key ('entry', `$key)" mode="text" >
				<xsl:with-param name="level" select="`$level" />
				<xsl:with-param name="back" select="true ()" />
			</xsl:apply-templates>
			<xsl:call-template name="track-down">
				<xsl:with-param name="level" select="`$level + 1" />
				<xsl:with-param name="chain" select="substring (`$chain, string-length (`$key) + 3)" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template match="atom:entry" mode="back">
		<xsl:param name="level" select="3" />
		<xsl:param name="loop" select="'[x]'" />
		<xsl:variable name="top" select="key ('back', atom:content/meta:properties/data:Id)" />
		<xsl:choose>
			<xsl:when test="contains(`$loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 461): Loop detected in chain </xsl:text><xsl:value-of select="concat('[', atom:content/meta:properties/data:Id, ']-', `$loop)" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="`$top">
						<xsl:apply-templates select="`$top" mode="back">
							<xsl:with-param name="level" select="`$level - 1" />
							<xsl:with-param name="loop" select="concat('[', atom:content/meta:properties/data:Id, ']-', `$loop)" />
						</xsl:apply-templates>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="track-down">
							<xsl:with-param name="chain" select="concat('[', atom:content/meta:properties/data:Id, ']-', `$loop)" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="atom:entry" mode="tags">
		<xsl:variable name="batt">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'BATT'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="tvvv">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'TVVV'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="weee">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'WEEE'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="scip">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'SCIP'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select="concat (atom:content/meta:properties/data:Material, ';', translate(atom:content/meta:properties/data:Description1, ';',' '), ';' , `$batt, ';', `$tvvv, ';', `$weee, ';', `$scip, '&#xD;&#xA;')" />
	</xsl:template>
	<xsl:template match="atom:entry" mode="duty">
		<xsl:variable name="hit">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="duty" select="`$SAP-DR-Product-Lookup.Duties [data:Duty = `$SAP-DR-Product-Lookup.Duty]/data:Id" />
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test="`$hit = 1">
			<xsl:value-of select="concat (atom:content/meta:properties/data:Material, ';', translate(atom:content/meta:properties/data:Description1, ';',' '), ';&#xD;&#xA;')" />
		</xsl:if>
	</xsl:template>
	<xsl:template match="/atom:feed">
		<xsl:choose>
			<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'attachments'">
				<xsl:value-of select="concat ('Base_Id', ';', 'Base_Material', ';', 'Base_Kurztext', ';', 'Item_Id', ';', 'Item_Material', ';', 'Item_Typ', ';', 'Item_Kurztext', '&#xA;')" />
				<xsl:apply-templates select="atom:entry [atom:content/meta:properties/data:Material = `$SAP-DR-Product-Lookup.Material]">
					<xsl:with-param name="level" select="0" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'forward'">
				<xsl:apply-templates select="atom:entry [`$SAP-DR-Product-Lookup.Material = '*' or atom:content/meta:properties/data:Material = `$SAP-DR-Product-Lookup.Material]">
					<xsl:with-param name="level" select="0" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'backward'">
				<xsl:apply-templates select="atom:entry[atom:content/meta:properties/data:Material = `$SAP-DR-Product-Lookup.Material]" mode="back" />
			</xsl:when>
			<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'tagexport'">
				<xsl:value-of select="'Material;Materialkurztext;BATT;TVVV;WEEE;SCIP&#xD;&#xA;'" />
				<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, `$SAP-DR-Product-Lookup.Product-ContentTypeID)]" mode="tags">
					<xsl:sort data-type="text" lang="en" order="ascending" select="atom:content/meta:properties/data:Material" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="`$SAP-DR-Product-Lookup.Mode = 'duty'">
				<xsl:variable name="tmp" select="`$SAP-DR-Product-Lookup.Duties [data:Duty = `$SAP-DR-Product-Lookup.Duty]" />
				<xsl:choose>
					<xsl:when test="count (`$tmp) = 1">
						<xsl:value-of select="'Material;Materialkurztext;&#xD;&#xA;'" />
						<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, `$SAP-DR-Product-Lookup.Product-ContentTypeID)]" mode="duty">
							<xsl:sort data-type="text" lang="en" order="ascending" select="atom:content/meta:properties/data:Material" />
						</xsl:apply-templates>
					</xsl:when>
					<xsl:otherwise>
						<xsl:message terminate="no">
							<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 567): Invalid duty &apos;</xsl:text><xsl:value-of select="`$SAP-DR-Product-Lookup.Duty" /><xsl:text>&apos;. Skipping.

</xsl:text>
						</xsl:message>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 578): Wrong mode &apos;</xsl:text><xsl:value-of select="`$SAP-DR-Product-Lookup.Mode" /><xsl:text>&apos;. Skipping.

</xsl:text>
				</xsl:message>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>

"@
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/SAP-DR-Product-Transfer.xslt (2023-01-05)
#
[Xml] $script:tnsxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	exclude-result-prefixes="atom com data meta">
	<xsl:output method="xml" encoding="utf-8" indent="yes" standalone="yes" />
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Transfer.ContentType" />
	<xsl:param name="SAP-DR-Product-Transfer.ProductLoadFile" />
	<xsl:param name="SAP-DR-Product-Transfer.ContentTypesLoadFile" />
	<xsl:variable name="SAP-DR-Product-Transfer.Products" select="document(`$SAP-DR-Product-Transfer.ProductLoadFile)/atom:feed/atom:entry" />
	<xsl:variable name="SAP-DR-Product-Transfer.ContentTypes" select="document(`$SAP-DR-Product-Transfer.ContentTypesLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry" />
	<xsl:key name="item" use="com:material" match="com:item" />
	<xsl:template match="data:element">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="pieces" select="`$pieces" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="data:element">
			<xsl:with-param name="pieces" select="`$pieces" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="pieces" select="`$pieces" />
			<xsl:with-param name="loop" select="`$loop" />
			<xsl:with-param name="base" select="`$base" />
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="atom:entry">
		<xsl:param name="pieces" />
		<xsl:param name="loop" select="'>>'" />
		<xsl:param name="base" select="." />
		<xsl:variable name="next" select="concat(`$loop, '-[', atom:content/meta:properties/data:Id, ']')" />
		<xsl:variable name="cotyid" select="atom:content/meta:properties/data:ContentTypeId" />
		<xsl:variable name="content" select="`$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name" />
		<xsl:choose>
			<xsl:when test="contains(`$loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Transfer.xslt (line 98): Loop detected in chain </xsl:text><xsl:value-of select="`$next" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="`$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name = `$SAP-DR-Product-Transfer.ContentType">
					<com:item>
						<com:material>
							<xsl:value-of select="atom:content/meta:properties/data:Material" />
						</com:material>
						<com:description>
							<xsl:value-of select="atom:content/meta:properties/data:Description1" />
						</com:description>
						<com:amount>
							<xsl:value-of select="`$pieces" />
						</com:amount>
					</com:item>
				</xsl:if>
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId">
					<xsl:with-param name="loop" select="`$next" />
					<xsl:with-param name="base" select="`$base" />
					<xsl:with-param name="pieces" select="`$pieces * atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId | atom:content/meta:properties/data:ItemListId">
					<xsl:with-param name="loop" select="`$next" />
					<xsl:with-param name="base" select="`$base" />
					<xsl:with-param name="pieces">
						<xsl:choose>
							<xsl:when test="`$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = `$cotyid]/data:Name = 'Product-in-Lots'">
								<xsl:value-of select="`$pieces div atom:content/meta:properties/data:Pieces" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="`$pieces" />
							</xsl:otherwise>
						</xsl:choose>					
					</xsl:with-param>
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="Object [Property/@Name='Material' and Property/@Name='Anzahl']">
			<xsl:apply-templates select="`$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:Material = current()/Property[@Name='Material']]">
				<xsl:with-param name="pieces" select="current()/Property[@Name='Anzahl']" />
			</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="Objects">
		<com:items>
			<xsl:apply-templates select="Object" />
		</com:items>
	</xsl:template>
	<xsl:template match="/com:items">
		<Root>
			<xsl:for-each select="com:item [generate-id (.) = generate-id (key ('item', com:material)[1])]">
				<xsl:sort select="com:material" />
				<Item>
					<Material>
						<xsl:value-of select="com:material" />
					</Material>
					<Materialkurztext>
						<xsl:value-of select="`$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:Material = current()/com:material]/atom:content/meta:properties/data:Description1" />
					</Materialkurztext>
					<Anzahl>
						<xsl:value-of select="sum (key ('item', com:material)/com:amount)" />
					</Anzahl>
				</Item>
			</xsl:for-each>
		</Root>
	</xsl:template>
</xsl:stylesheet>
"@
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare/Abfallwirtschaft/XSLT/Waste-Batch-Countries.xslt (2024-02-05)
#
[Xml] $script:wbcxsl = @"
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:xs="http://www.w3.org/2001/XMLSchema" 
	xmlns:fn="http://www.w3.org/2005/xpath-functions" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons"
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty fn prd xs">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone="yes" />
	<xsl:key name="countries" match="dty:references/dty:entry/com:countries/com:element" use="." />
	<xsl:template match="/prd:root">
		<root>
			<options default="1">
				<prompt>Select country:</prompt>
				<xsl:for-each select="dty:references/dty:entry/com:countries/com:element [generate-id (.) = generate-id (key('countries',.)[1])]">
					<xsl:sort select="." />
					<option id="{position()}">
						<key><xsl:value-of select="." /></key>
						<label><xsl:value-of select="." /></label>
					</option>
				</xsl:for-each>
			</options>
		</root>
	</xsl:template>
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#	XML Formulare\Abfallwirtschaft\XSLT\Waste-Batch-Country-Products.xslt (2024-03-22)
#
[Xml] $script:wbcpxsl = @"
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:xs="http://www.w3.org/2001/XMLSchema" 
	xmlns:fn="http://www.w3.org/2005/xpath-functions" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons"
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty fn prd xs">
	<xsl:param name="Global.country" select="'BE'" />
	<xsl:output method="text" encoding="UTF-8" />
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<xsl:key name="duties" match="dty:references/dty:entry" use="@SPKey"/>
	<xsl:key name="munch" match="prd:product/prd:duties/prd:duty" use="concat (../../@SPKey, '-', @SPKey)" />
	<xsl:key name="countries" match="dty:references/dty:entry/com:countries/com:element" use="." />
	<xsl:template match="prd:product" mode="SCIP">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="`$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat (`$this/prd:material, ';', key ('duties', @SPKey)[1]/dty:code, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count (`$current-group), '#.##0'),'&#xA;')"/>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="prd:product" mode="TVVV">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="`$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat (`$this/prd:material, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count (`$current-group), '#.##0'), ';', format-number (sum (`$current-group/prd:data/VV_x002d_Alu), '#.##0'), ';', format-number (sum (`$current-group/prd:data/VV_x002d_Steel) + sum (`$current-group/prd:data/VV_x002d_Tinplate), '#.##0'), ';', format-number (sum (`$current-group/prd:data/Paper), '#.##0'), ';', format-number (sum (`$current-group/prd:data/VV_x002d_Plastic), '#.##0'), '&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="prd:product" mode="BATT-WEEE">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="`$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = `$Global.country][`$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat (`$this/prd:material, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count (`$current-group), '#.##0'), ';', format-number(sum (`$current-group/prd:data/Weight), '#.##0,000'), '&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="/prd:root">
		<xsl:variable name="batch" select="@batch" />
		<list>
			<xsl:choose>
				<xsl:when test="`$batch = 'BATT' or `$batch = 'WEEE'">
					<xsl:text>material; category; number of parts; total weight of parts [kg]&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = `$Global.country and dty:batch = `$batch]/@SPKey]" mode="BATT-WEEE">
						<xsl:with-param name="batch" select="`$batch" />
						<xsl:sort data-type="text" select="prd:material" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:when test="`$batch = 'TVVV'">
					<xsl:text>material; category; number of parts; aluminum [gr]; steel/tinplate (FE 04) [gr]; paper [gr]; plastic [gr]&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = `$Global.country and dty:batch = `$batch]/@SPKey]" mode="TVVV">
						<xsl:with-param name="batch" select="`$batch" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:when test="`$batch = 'SCIP'">
					<xsl:text>material; category; number of parts&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = `$Global.country and dty:batch = `$batch]/@SPKey]" mode="SCIP">
						<xsl:with-param name="batch" select="`$batch" />
						<xsl:sort data-type="text" select="prd:material" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:otherwise>
					<xsl:message>
						<xsl:text>[ERROR] 'Batch' must be either WEEE, BATT or TVVV</xsl:text>
					</xsl:message>
				</xsl:otherwise>
			</xsl:choose>
		</list>
	</xsl:template>
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare\Common\ps1\msxml.ps1 (2022-12-22)
#		
function script:msxml([Object] $xsl, [Object] $xml, [String] $out = '', [System.Collections.Hashtable] $param = @{}) {
	#
	[System.Xml.XmlUrlResolver] $local:xres = New-Object -typeName System.Xml.XmlUrlResolver
	[System.Xml.Xsl.XsltSettings] $local:xset = New-Object -typeName System.Xml.Xsl.XsltSettings
	#
	$null = $xset.set_EnableDocumentFunction($true)
	$null = $xset.set_EnableScript($true)
	#
	[System.Xml.Xsl.XslCompiledTransform] $script:xslt = New-Object -typeName System.Xml.Xsl.XslCompiledTransform
	[System.Xml.Xsl.XsltArgumentList] $local:xal = New-Object -typename System.Xml.Xsl.XsltArgumentList
	#
	if($verbose -eq $true) {
		$xal.AddParam("verbose", "", 1)
	} else {
		$xal.AddParam("verbose", "", 0)
	}
	if($debug -eq $true) {
		$xal.AddParam("debug", "", 1)
	} else {
		$xal.AddParam("debug", "", 0)
	}
	#
	if($null -ne $param) {
		[String[]] $local:val = $null
		foreach($val in $param.get_Keys()) {
			Write-Debug @"
[DEBUG] msxml::msxml(...): Adding XSLT parameter $val = $($param[$val])
"@
			$xal.AddParam($val, "", $($param[$val]))
		}
	}
	#
	if($xsl.GetType().Name -eq 'XmlDocument') {
		#
		$xslt.Load([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xsl.PsBase.InnerXml)), $local:xset, $local:xres)
		#
	} elseif ([System.IO.File]::Exists($xsl) -eq $true) {
		#
		$xslt.Load($xsl, $local:xset, $local:xres)
		#
	} else {
		Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed.
"@
	}
	#
	if ($out -ne '') {
		#
		[System.IO.FileStream] $local:fsm = New-Object System.IO.FileStream -ArgumentList @($out, 2)
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $local:xal, $local:fsm)
			#
		} elseif ([System.IO.File]::Exists($xml) -eq $true) {
			#
			$xslt.Transform($xml, $local:xal, $local:fsm)
			#
		} else {
			Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed.
"@
		}
		#
		$local:fsm.Close()
		$local:fsm.Dispose()
		#
		return $null
		#
	} else {
		#	
		[System.IO.StringWriter] $local:osw = New-Object System.IO.StringWriter
		#
		if($xml.GetType().Name -eq 'XmlDocument') {
			#
			$xslt.Transform([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $xml.PsBase.InnerXml)), $xal, $local:osw)
			#
		} elseif ([System.IO.File]::Exists($xml) -eq $true) {
			#
			$xslt.Transform($xml, $xal, $local:osw)
			#
		} else {
			Write-Verbose @"
[FATAL] msxml::msxml(...): Transformation failed.
"@
		}
		$local:osw.close()
		[String] $local:res = $local:osw.ToString()
		$local:osw.Dispose()
		#
		return $local:res
	}
}
#
# -----------------------------------------------------------------------------------------------
#
[Xml] $local:men = @"
<?xml version="1.0" encoding="UTF-8"?>
<root>
	<options default="1">
		<prompt>Which action do you wish to perform?</prompt>
		<option id="1">
			<key>BATT</key>
			<label>Consolidate battery waste amount according to recycling duties</label>
		</option>
		<option id="2">
			<key>WEEE</key>
			<label>Consolidate electric waste amount according to recycling duties</label>
		</option>
		<option id="3">
			<key>TVVV</key>
			<label>Consolidate packaging waste amount according to recycling duties</label>
		</option>
		<option id="4">
			<key>CHECK</key>
			<label>Display the electric, battery, packaging items referenced by catalogue numbers ...</label>
		</option>
		<option id="5">
			<key>BACK</key>
			<label>Trace back electric, battery, packaging item to referencing catalogue numbers ...</label>
		</option>
		<option id="6">
		<key>DUTY</key>
			<label>Lookup catalogue products linked to a specific recycling duty ...</label>
		</option>
		<option id="7">
			<key>IMAGES</key>
			<label>Download attached images and files ...</label>
		</option>
		<option id="8">
			<key>TRANSFER</key>
			<label>Calculate bucket transfer from cataloque products into packaging, electrics or batteries ...</label>
		</option>
		<option id="9">
			<key>NOTES</key>
			<label>Export comments into a .csv UTF-8 spreadsheet ...</label>
		</option>
		<option id="10">
			<key>TAGS</key>
			<label>Export recycling tags [BATT,TVVV,WEEE] into a .csv UTF-8 spreadsheet ...</label>
		</option>
		<option id="11">
			<key>EXPORT</key>
			<label>Export product-by-duty [TVVV/WEEE/BATT] into a .csv UTF-8 spreadsheet ...</label>
		</option>
		<option id="12">
			<key>RELOAD</key>
			<label>Set master data reload flag (master data will be reloaded in next step)</label>
		</option>
		<option id="13">
			<key>BEARER</key>
			<label>Save the bearer token to disk ...</label>
		</option>
	</options>
</root>
"@
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare\Common\ps1\GetOptMap.ps1 (2018-11-29)
#
function script:getOptMap([System.Xml.XmlElement] $opts, [Int] $max, [String] $brk = "x", [Int] $col = 1) {
	#
	[String] $local:title = "Which action do you wish to perform?"
	[System.Collections.Hashtable] $local:hsh = @{}
	#
	if($null -ne $opts.options.prompt) {
		$title = $opts.options.prompt
	}
	#
	[System.Text.StringBuilder] $local:msg = new-object -typeName System.Text.StringBuilder -argumentList @"
	
	$title

"@
	#
	foreach($local:pos in $opts.options.option) {
		$hsh[[Int]$pos.id] = ($pos.key, $pos.label)
		[String] $local:fill = " " * (7 - $("$($pos.id)".Length))
		$null = $msg.append(@"
	
	[$($pos.id)]$local:fill$($hsh[[Int]$pos.id][1])
"@)
	}
	#
	$null = $msg.append(@"
	
	
	[$brk]$(" " * 6)Exit the script and remove temporary files

"@)
	[String] $trailer = @"

Please enter the number given in square brackets [#] according to your selection
"@
	#
	[System.Collections.Hashtable] $local:res = @{}
	[String] $local:lasterr = ""
	#
	do {
		#
		[String] $local:inval = ""
		[String] $local:label = $msg.append($lasterr).append($trailer)
		[Boolean] $local:end = $true
		#
		do {
			#
			$inval = $(Read-Host -prompt $label)
			#
		} while ($inval -eq "")
		#
		$null = $res.Clear()
		$lasterr = ""
		#
		foreach($local:item in $inval.split((","))) {
			if ($item -eq $brk) {
				return @{}
			} else {
				if ($null -ne $item -as [Int]) {
					[Int] $local:id = [Int] $item
					if($hsh.contains($id)) {
						if($res.contains($id) -eq $false) {
							$res[$id] = $hsh[$id][0]
						}
					} else {
						$lasterr = @"
$lasterr
						
	***	ERROR: The number $id is not included in the list of valid choices.	***

"@
						$end = $false
					}
				} else {
					$lasterr = @"
$lasterr
					
	***	ERROR: The input value $item is not a number as required.	***

"@
					$end = $false
				}
			}
		}
		#
		if ($res.get_Count() -gt $max) {
			$lasterr = @"
$lasterr
			
	***	ERROR: Too many values have been selected. Maximum is $max.	***

"@
			$end = $false
		}
		#
	} until ($end -eq $true)
	#
	return $res
	#
}
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare\Common\ps1\OpenFileDialog.ps1 (2022-10-01)
#
function script:OpenFileDialog([String] $title = "Oups.", [String] $type = "all", [String] $defpath = "") {
	#
	[System.Windows.Forms.OpenFileDialog] $openFileDialog1 = new-object -typeName System.Windows.Forms.OpenFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'xmlcsv'='Excel 2003 XML or Excel CSV file (*.xml)|*.xml;*csv';'csv'='Comma separated values spread sheet (*.csv)|*.csv'; 'htm'='HTML source (*.htm, *.html)|*.htm;*.html'; 'i6z'='IUCLID (*.i6z)|*.i6z'; 'all'='All files (*.*)|*.*'}
	#
	[String] $local:res = $null
	#
	$openFileDialog1.Reset()
	#
	if([System.IO.Directory]::Exists($defpath)) {
		Write-Debug @"
[Debug] SAP-EAR-DSD-Converter::OpenFileDialog: Default path set to $defpath
"@
		$openFileDialog1.set_InitialDirectory($defpath)
	} else {
		Write-Debug @"
[Debug] SAP-EAR-DSD-Converter::OpenFileDialog: Default path not set.
"@
	}
	$openFileDialog1.set_Filter($filters[$type])
	$openFileDialog1.set_FilterIndex(2)
	$openFileDialog1.set_Title($title)
	$openFileDialog1.set_AddExtension($true)
	$openFileDialog1.set_AutoUpgradeEnabled($true)
	$openFileDialog1.set_CheckFileExists($true)
	$openFileDialog1.set_CheckPathExists($true)
	$openFileDialog1.set_ShowHelp($false)
	$openFileDialog1.set_RestoreDirectory($true)
	$openFileDialog1.set_Multiselect($false)
	$openFileDialog1.set_ShowReadOnly($false)
	#
	if($openFileDialog1.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$res = $openFileDialog1.get_FileName()
        } else {
		$res = $null
	}
	#
	$openFileDialog1.Dispose()
	#
	return $res
}
#
# -----------------------------------------------------------------------------------------------
#
#	Originaldatei (Version):
#		XML Formulare\Common\ps1\SaveFileDialog.ps1 (2022-10-01)
#
function script:SaveFileDialog([String] $title = "Write file", [String] $type = 'csv', [String] $defpath = "", [String] $defname = $null, [Boolean] $delete = $false) {
	#
	[System.Windows.Forms.SaveFileDialog] $local:saveFileDialog1 = new-object -typeName System.Windows.Forms.SaveFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'csv'='CSV spreadsheet (*.csv)|*.csv' ; 'xml'='XML file (*.xml)|*.xml'; 'xsl' = 'XSL transformation (*.xsl, *.xslt)|*.xsl,*.xslt'; 'all'='Alle Dateien (*.*)|*.*'}
	#
	[String] $local:res = $null
	#
	$saveFileDialog1.Reset()
	$saveFileDialog1.set_InitialDirectory($defpath)
	$saveFileDialog1.set_Filter($filters[$type])
	$saveFileDialog1.set_FilterIndex(2)
	$saveFileDialog1.set_Title($title)
	$saveFileDialog1.set_AddExtension($true)
	$saveFileDialog1.set_AutoUpgradeEnabled($true)
	$saveFileDialog1.set_CheckPathExists($true)
	$saveFileDialog1.set_ShowHelp($false)
	$saveFileDialog1.set_RestoreDirectory($true)
	$saveFileDialog1.set_OverwritePrompt($true)
	if($null -ne $defname) {
		$saveFileDialog1.set_FileName($defname)
	}
	#
	if($saveFileDialog1.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$res = $saveFileDialog1.get_FileName()
		#
		if ([System.IO.File]::Exists($res) -eq $true -and $delete -eq $true) {
			Remove-Item -force $res
		}
		
        } else {
		$res = $null
	}
	#
	$saveFileDialog1.Dispose()
	#
	return $res
}
#
function script:CompileMaster([String] $cmBatch) {
	#
	[System.Collections.Hashtable] $private:prm = @{'SAP-DR-Product-Duties.DutyLoadFile'="$script:tmpDutiesXml" ; 'SAP-DR-Product-Duties.Batch'="$cmBatch"; 'SAP-DR-Product-Duties.Product-ContentTypeID'='0x01003FAF714C6769BF4FA1B36DCF47ED659702' }
	#
	. script:msxml -xsl $script:pdmxsl -xml $script:tmpProductsXml -out $script:tmpMasterXml -param $private:prm
	#
	if ($debug) {
		#
		$null = Read-Host -Prompt @"

Master data has been compiled for batch '$bat' and written to:

- Batchdata: . '$script:tmpMasterXml'

Hit ENTER to continue ...
"@
		#
	}
	#
	if([System.IO.Directory]::Exists("$copies")) {
		#
		Copy-Item -Force -Path $script:tmpMasterXml -Destination "$copies\master.xml"
		#
	}
	#
}
# -----------------------------------------------------------------------------------------------
#
#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
#	und entsprechend den Parametern $verbose und $debug neu setzen.
#
[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
#
if($verbose -eq $true -or  $debug -eq $true) {
	#
	Write-Verbose @"
SAP-DR-Reporting.ps1::main(...): switched to verbose mode.
"@
	#
	$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
	#
	if($debug -eq $true) {
		Write-Verbose @"
SAP-DR-Reporting.ps1::main(...): switched to debug mode.
"@
		$DebugPreference = [System.Management.Automation.ActionPreference]::Inquire
	} else {
		$DebugPreference = [System.Management.Automation.ActionPreference]::Continue
	}
} else {
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
}
#
# -----------------------------------------------------------------------------------------------
#
$private:passphrase = {
	#
	if ($env:__Waste_Reporting_Key -ne $null) {
		#
		return $($env:__Waste_Reporting_Key)
		#
	} else {
		#
		return $(Read-Host -Prompt 'Please enter password')
		#
	}
}
#
[String] $private:act = $(. local:getAccessToken -phrase $(. $private:passphrase))
#[String] $private:act = $(Get-Content "$env:HOMEDRIVE\$env:HOMEPATH\Desktop\token.txt")
#
[String] $local:tmp1 = [System.IO.Path]::GetTempFileName()
[String] $local:tmp2 = [System.IO.Path]::GetTempFileName()
[String] $local:tmp3 = [System.IO.Path]::GetTempFileName()
[String] $local:tmp4 = [System.IO.Path]::GetTempFileName()
#
[String] $script:tmpTypesXml = "$($local:tmp1 -replace '\.tmp$', "-$pid.xml")"
[String] $script:tmpFieldsXml = "$($local:tmp2 -replace '\.tmp$', "-$pid.xml")"
[String] $script:tmpDutiesXml = "$($local:tmp3 -replace '\.tmp$', "-$pid.xml")"
[String] $script:tmpProductsXml = "$($local:tmp4 -replace '\.tmp$', "-$pid.xml")"
#
Rename-Item $local:tmp1 $script:tmpTypesXml
Rename-Item $local:tmp2 $script:tmpFieldsXml
Rename-Item $local:tmp3 $script:tmpDutiesXml
Rename-Item $local:tmp4 $script:tmpProductsXml
#
. local:getListItems -mod 'types' -act $act -out $script:tmpTypesXml
#
. local:getListItems -mod 'fields' -act $act -out $script:tmpFieldsXml
#
if ($debug) {
	#
	$null = Read-Host -Prompt @"
    
    Descriptive data has been fetched from SharePoint into temporary files:

        - Fields: ........ '$script:tmpFieldsXml'
        - ContentTypes: .. '$script:tmpTypesXml'

    Hit ENTER to continue ...
"@
}
#
if([System.IO.Directory]::Exists("$copies")) {
	#
    Copy-Item -Force -Path $script:tmpFieldsXml -Destination "$copies\fields.xml"
	Copy-Item -Force -Path $script:tmpTypesXml -Destination "$copies\types.xml"
	#
}
#
[Boolean] $local:reload = $true
#
while ($private:act.Length -gt 1) {
	#
	if (($verbose -or $debug) -ne $true) { Clear-Host }
	#
	[System.Collections.IEnumerator] $local:cho = $(. script:getOptMap -opts ($local:men).root -max 1).get_Values().GetEnumerator()
	#
	if ($local:cho.moveNext() -eq $false) {
		#
		break ;
		#
	}
	#
	[String] $private:bat = $local:cho.get_Current()
	#
	[String] $local:tmp5 = [System.IO.Path]::GetTempFileName()
	[String] $local:tmp6 = [System.IO.Path]::GetTempFileName()
	#
	[String] $script:tmpMasterXml = "$($local:tmp5 -replace '\.tmp$', "-$pid.xml")"
	[String] $script:tmpInterXml = "$($local:tmp6 -replace '\.tmp$', "-$pid.xml")"
	#
	Rename-Item $local:tmp5 $script:tmpMasterXml
	Rename-Item $local:tmp6 $script:tmpInterXml
	#
	if ($local:reload) {
		#
		. local:getListItems -mod 'products' -act $act -out $script:tmpProductsXml
		#
		. local:getListItems -mod 'duties' -act $act -out $script:tmpDutiesXml
		#
		$local:reload = $false
		#
		if([System.IO.Directory]::Exists("$copies")) {
			#
			Copy-Item -Force -Path $script:tmpDutiesXml -Destination "$copies\duties.xml"
			Copy-Item -Force -Path $script:tmpProductsXml -Destination "$copies\products.xml"
			#
		}
		#
	}
	#
	if ($debug) {
		#
		$null = Read-Host -Prompt @"
    
    Master data has been fetched from SharePoint into temporary files:

        - Duties: .... '$script:tmpDutiesXml'
        - Products: .. '$script:tmpProductsXml'

	Press ENTER to continue ...
"@
	}
	#
	switch ($private:bat)
	{
        #
		'CHECK' {
			#
			if (($verbose -or $debug) -ne $true) { Clear-Host }
			#
			[String] $local:mat = Read-Host -Prompt @"
++			
++	Enter a catalogue number as root product to trace down
"@
			#
			while ($local:mat.length -gt 0) {
				#
				Write-Host @"
++
"@
				#
				[System.Collections.Hashtable] $local:prm = @{'SAP-DR-Product-Lookup.Mode'='forward' ; 'SAP-DR-Product-Lookup.Material'="$local:mat"; 'SAP-DR-Product-Lookup.FieldsLoadFile'="$script:tmpFieldsXml" ; 'SAP-DR-Product-Lookup.ContentTypesLoadFile'="$script:tmpTypesXml"; 'SAP-DR-Product-Lookup.DutyLoadFile'="$script:tmpDutiesXml"; 'SAP-DR-Product-Lookup.ShowDuties'='1' }  # 'SAP-DR-Product-Lookup.UnfoldBundles'='1' entfaltet gebuendelte Komponenten auf die festgelgte Anzahl
				#
				. script:msxml -xsl $script:lupxsl -xml $script:tmpProductsXml -param $local:prm
				#
				$local:mat = Read-Host -Prompt @"
++				
++	Enter another product catalogue number to process or press ENTER to return to start menue
"@
				#
			} 
		}
		#
		'BACK' {
			#
			if (($verbose -or $debug) -ne $true) { Clear-Host }
			#
			[String] $local:mat = Read-Host -Prompt @"
++			
++	Enter a component material identifier as the leaf for tracing linkage back
"@
			#
			while ($local:mat.length -gt 0) {
				#
				Write-Host @"
++
"@
				#
				[System.Collections.Hashtable] $local:prm = @{'SAP-DR-Product-Lookup.Mode'='backward' ; 'SAP-DR-Product-Lookup.Material'="$local:mat"; 'SAP-DR-Product-Lookup.FieldsLoadFile'="$script:tmpFieldsXml" ; 'SAP-DR-Product-Lookup.ContentTypesLoadFile'="$script:tmpTypesXml"; 'SAP-DR-Product-Lookup.DutyLoadFile'="$script:tmpDutiesXml" } 
				#
				. script:msxml -xsl $script:lupxsl -xml $script:tmpProductsXml -param $local:prm
				#
				$local:mat = Read-Host -Prompt @"
++				
++	Enter another catalogue number to process or press ENTER to return to start menue
"@
				#
			} 
		}
		#
		'TAGS' {
			#
			[System.Collections.Hashtable] $local:prm = @{'SAP-DR-Product-Lookup.Mode'='tagexport' ; 'SAP-DR-Product-Lookup.FieldsLoadFile'="$script:tmpFieldsXml" ; 'SAP-DR-Product-Lookup.ContentTypesLoadFile'="$script:tmpTypesXml"; 'SAP-DR-Product-Lookup.DutyLoadFile'="$script:tmpDutiesXml"; 'SAP-DR-Product-Lookup.Product-ContentTypeID'='0x01003FAF714C6769BF4FA1B36DCF47ED659702' }
			#
			[String] $private:text = $(. script:msxml -xsl $script:lupxsl -xml $script:tmpProductsXml -param $local:prm)
			#
			ConvertFrom-Csv -Delimiter ';' -InputObject $private:text | Out-GridView -Wait 
			#
			[String] $private:save = $(. script:SaveFileDialog -defpath $pwd -defname "recytags.csv")
			#
			if($null -ne $private:save -and $private:save -ne "") {
				#
				Out-File -InputObject $private:text -Encoding utf8 -FilePath $private:save
				#
			}
			#
		}
		'DUTY'	{
			#
			[String] $local:duty = Read-Host -Prompt @"
++			
++	Enter recycling duty to trace into catalogue products
"@
			#
			[System.Collections.Hashtable] $local:prm = @{'SAP-DR-Product-Lookup.Mode'='duty' ; 'SAP-DR-Product-Lookup.Duty'="$local:duty" ; 'SAP-DR-Product-Lookup.FieldsLoadFile'="$script:tmpFieldsXml" ; 'SAP-DR-Product-Lookup.ContentTypesLoadFile'="$script:tmpTypesXml"; 'SAP-DR-Product-Lookup.DutyLoadFile'="$script:tmpDutiesXml"; 'SAP-DR-Product-Lookup.Product-ContentTypeID'='0x01003FAF714C6769BF4FA1B36DCF47ED659702' }
			#
			[String] $private:text = $(. script:msxml -xsl $script:lupxsl -xml $script:tmpProductsXml -param $local:prm)
			#
			ConvertFrom-Csv -Delimiter ';' -InputObject $private:text | Out-GridView -Wait
			#
		}
		#
		'BEARER' {
			#	
			[String] $private:save = $(. script:SaveFileDialog -defpath $pwd -defname "token.txt")
			#
			if($null -ne $private:save -and $private:save -ne "") {
				#
				Out-File -InputObject $private:act -Encoding utf8 -FilePath $private:save
				#
			}
			#
		}
		#
		'NOTES'	{
			#
			. local:getComments | ConvertFrom-Csv -Delimiter ";" | Out-GridView -Title "Comments in $local:list - select rows to include in spreadsheet" -PassThru | Export-Csv -Delimiter ";" -Encoding UTF8 -Path $(. script:SaveFileDialog -defpath $pwd -defname "comments.csv") -NoTypeInformation
			#
		}
		#
		# Die Alternative Transfer berechnet den Transfer eines Vektors aus Produktzahlen auf einen Vektor aus Verpackungszahlen
		# Der Produktvektor muss als .csv Datei mit zwei Spalten geladen werden. Die erste Zeile enth"alt Spalten"uberschriften:
		#
		# Material;Anzahl
		# 1000002322;23
		# ...;...
		#
		'TRANSFER' {
			#
			[String] $local:load = $(. script:OpenFileDialog -title "Datei mit Produktvektor auswaehlen" -type "csv" -defpath $DataDir)
            #
			if ([System.IO.File]::Exists($local:load) -eq $true) {
                #
                [Xml] $local:menue = @"
<?xml version="1.0" encoding="UTF-8"?>
<root>
	<options default="1">
		<prompt>Which transfer do you want to calculate?</prompt>
		<option id="1">
			<key>Electric</key>
			<label>Electric and electronic devices except of batteries</label>
		</option>
		<option id="2">
			<key>Battery</key>
			<label>Batteries and battery modules</label>
		</option>
		<option id="3">
			<key>Packaging</key>
			<label>Packaging and parts thereof</label>
		</option>
    </options>
</root>
"@
                #
	            [System.Collections.IEnumerator] $local:choice = $(. script:getOptMap -opts ($local:menue).root -max 1).get_Values().GetEnumerator()
	            #
	            if ($local:choice.moveNext() -eq $true) {
				    #
				    [Xml] $local:prods = $(Import-Csv -Delimiter ";" -Path $local:load | ConvertTo-Xml)
				    #
				    [Xml] $local:inter = $(. script:msxml -xsl $script:tnsxsl -xml $local:prods -param @{'SAP-DR-Product-Transfer.ContentType'="$($local:choice.get_Current())"; 'SAP-DR-Product-Transfer.ProductLoadFile'="$script:tmpProductsXml"; 'SAP-DR-Product-Transfer.ContentTypesLoadFile'="$script:tmpTypesXml"})
				    #
				    [Xml] $local:trans = $(. script:msxml -xsl $script:tnsxsl -xml $local:inter -param @{'SAP-DR-Product-Transfer.ProductLoadFile'="$script:tmpProductsXml"; 'SAP-DR-Product-Transfer.ContentTypesLoadFile'="$script:tmpTypesXml"})
				    #
				    $($local:trans).Root.ChildNodes | Select-Object -Property Material,Materialkurztext,@{Name="Amount"; Expression={$_.Anzahl -as [Int]}} | Out-GridView -Title "Products transferred to $($local:choice.get_Current())" -PassThru | Export-Csv -Path $(. script:SaveFileDialog -defpath $pwd -defname "$($local:choice.get_Current()).csv") -Delimiter:";" -Encoding:utf8 -NoTypeInformation
				    #
			    }
			#
            }
            #
		}
		#
		#	'EXPORT' exports a colon separated UTF-8 encoded list of product 'Material' keys 
		#	grouped by their respective attached duty with quantity and qualitative master data
		#
		'EXPORT' {
			#
			[Xml] $private:menue = @"
<?xml version="1.0" encoding="UTF-8"?>
<root>
	<options default="1">
		<prompt>Select the batch:</prompt>
		<option id="1">
			<key>WEEE</key>
			<label>Electric and electronic devices except of BATT</label>
		</option>
		<option id="2">
			<key>BATT</key>
			<label>Batteries and battery modules</label>
		</option>
		<option id="3">
			<key>TVVV</key>
			<label>Packaging for primary, secondary and transport</label>
		</option>
	</options>
</root>
"@
			#
			[System.Collections.IEnumerator] $local:cbat = $(. script:getOptMap -opts ($private:menue).root -max 1).get_Values().GetEnumerator()
			#
			if ($local:cbat.moveNext() -eq $true) {
				#
				. script:CompileMaster -cmBatch $($local:cbat.get_Current())
				#
				[Xml] $local:countries = $(. script:msxml -xsl $script:wbcxsl -xml $script:tmpMasterXml)
				#
				[System.Collections.IEnumerator] $local:ccon = $(. script:getOptMap -opts ($local:countries).root -max 1).get_Values().GetEnumerator()
				#
				if ($local:ccon.moveNext() -eq $true) {
					#
					[String] $local:result = $(. script:msxml -xsl $script:wbcpxsl -xml $script:tmpMasterXml -param @{'Global.country'="$($local:ccon.get_Current())"})
					#
					[String] $private:save = $(. script:SaveFileDialog -defpath $pwd -defname "product-duties-$($local:ccon.get_Current())-$($local:cbat.get_Current()).csv")
					#
					if($null -ne $private:save -and $private:save -ne "") {
						#
						Out-File -InputObject $local:result -Encoding utf8 -FilePath $private:save
						#
					} 
					#
				}
				#
			}
			#
		}
		#	'IMAGES' downloads all images from the entire tree of items starting at and 
		#	including the item input 'Material' key. The function 
		#
		'IMAGES' {
			#
			if (($verbose -or $debug) -ne $true) { Clear-Host }
			#
			[System.Windows.Forms.FolderBrowserDialog] $private:folderBrowserDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
            #
			$private:folderBrowserDialog.set_SelectedPath($DataDir)
			#
			if ($private:folderBrowserDialog.ShowDialog() -eq 'OK') {
				#
				[String] $private:folder = $private:folderBrowserDialog.get_SelectedPath()
				#
				[String] $private:tmpImage = [System.IO.Path]::GetTempFileName()
				#
				[String] $private:mat = Read-Host -Prompt @"
|			
|	Enter catalogue number for download of attached images and files or press ENTER to return to start menue
"@
				#
				while ($private:mat.length -gt 0) {
					#
					Write-Host @"
|
"@
					#
					[System.Collections.Hashtable] $private:prm = @{'SAP-DR-Product-Lookup.Mode'='attachments' ; 'SAP-DR-Product-Lookup.Material'="$local:mat"; 'SAP-DR-Product-Lookup.FieldsLoadFile'="$script:tmpFieldsXml" ; 'SAP-DR-Product-Lookup.ContentTypesLoadFile'="$script:tmpTypesXml"; 'SAP-DR-Product-Lookup.DutyLoadFile'="$script:tmpDutiesXml"; 'SAP-DR-Product-Lookup.ShowDuties'='1' } 
					#
					foreach ($row in $(. script:msxml -xsl $script:lupxsl -xml $script:tmpProductsXml -param $private:prm | ConvertFrom-Csv -Delimiter ';')) {
						#
						foreach ($decoRowVal in $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/json"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/Items($($row.Item_Id))/AttachmentFiles").value) {
							#
							[Object] $private:props = $(Invoke-RestMethod -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/json"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/Properties")
							#
							[String] $private:pattern = "^(.*?)\.(bmp|BMP|gif|GIF|png|PNG|jpg|JPG|jpeg|JPEG)$"
							#
							if ($decoRowVal.FileName -match $private:pattern) {
								#
								[String] $private:outFile = "$private:folder\$($row.Base_Material)_$($row.Item_Typ)_$($row.Item_Id)_$($($decoRowVal.FileName) -replace $private:pattern,'$1').png"
								#
								Write-Host  @"
			
SAP-DR-Reporting.ps1::main (): Downloading image '$($decoRowVal.FileName)' 
|	device: ...... $($private:props.wic_x005f_System_x005f_Photo_x005f_CameraManufacturer) $($private:props.wic_x005f_System_x005f_Photo_x005f_CameraModel)
|	created: ..... $($private:props.vti_x005f_timelastwritten)
|	disk size: ... $($private:props.vti_x005f_filesize) bytes
|	writing to: . $private:outfile
"@
								#
								Invoke-WebRequest -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="image/jpeg, image/png, image/gif, image/pjpeg"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/`$value" -OutFile $private:tmpImage
								#
								if ($($row.Item_Id) -ne $($row.Base_Id)) {
									#
									. local:decorate -decoInFile $private:tmpImage -decoOutFile $private:outFile -headingL "$($row.Base_Material) $($row.Base_Kurztext)" -headingR $($private:props.vti_x005f_timelastwritten) -subheading "$($row.Item_Typ)($($row.Item_material)): $($row.Item_Kurztext)"
									#
								} else {
									#
									. local:decorate -decoInFile $private:tmpImage -decoOutFile $private:outFile -headingL "$($row.Base_Material) $($row.Base_Kurztext)" -headingR $($private:props.vti_x005f_timelastwritten)
									#
								}
								#
							} else {
								#
								[String] $private:outFile = "$private:folder\$($row.Base_Material)_$($row.Item_Typ)_$($row.Item_Id)_$($decoRowVal.FileName)"
								#
								Write-Host  @"
			
SAP-DR-Reporting.ps1::main (): Downloading file '$($decoRowVal.FileName)'
|	created: ..... $($private:props.vti_x005f_timelastwritten)
|	disk size: ... $($private:props.vti_x005f_filesize) bytes
|	writing to: . $private:outfile
"@
								#
								Invoke-WebRequest -Method Get -Headers @{"Authorization"="Bearer $local:act";"Accept"="application/octet-stream"} -Uri "https://$tenant.sharepoint.com/sites/$site/_api/web/getFileByServerRelativeUrl('$($decoRowVal.ServerRelativeUrl)')/`$value" -OutFile $private:outFile
								#
							}
						}
					}
					#
					$private:mat = Read-Host -Prompt @"
				
|	Enter another catalogue number or press ENTER to return to start menue
"@
				#
				}
				#
				local:cleanup -loc $private:tmpImage
				#
			}
		}
		#
		# 'RELOAD'
		#
		'RELOAD' {
			#
			$local:reload = $true
			#
		}
		#
		default {
			#
			. script:CompileMaster -cmBatch $private:bat
			#
			[String] $local:load = $(. script:OpenFileDialog -title "Select file with $private:bat export data (Excel 2003 XML oder UTF-8 CSV)" -type "xmlcsv" -defpath $DataDir)
			#
			if ([System.IO.File]::Exists($local:load) -eq $true) {
				#
				#	Das Verzeichnis der Auswahl als Startpunkt fuer die Suche nach dem
				#	Ablageort der konvertierten Datei und fuer die folgende Runde festlegen.
				#
				$DataDir = $(Get-Item $local:load).DirectoryName
				#
                [Object] $local:data = $null
                #
                if ($local:load -match '\.xml$|\.XML$') {
                    #
                    $local:data = $local:load
                    #
                } elseif ($local:load -match '\.csv$|\.CSV$') {
                    #
                    $local:data = $($(. msxml -xsl $script:intxsl -xml $(Import-Csv -Encoding utf8 -Delimiter ";" -Path $local:load | ConvertTo-Xml) -param @{"SAP-DR-Recycling-Preprocess.NumberGroupSeparator"="$((Get-Culture).NumberFormat.NumberGroupSeparator)"; "SAP-DR-Recycling-Preprocess.NumberDecimalSeparator"="$((Get-Culture).NumberFormat.NumberDecimalSeparator)"}) -as [Xml])
                    #
                } else {
                    #
                    break
                    #
                }
                #
				[Xml] $private:mea = $(. script:msxml -xsl $script:ccsxsl -xml $local:data -param @{"SAP-DR-Recycling-Countries.MasterData"="$script:tmpMasterXml"})
				#
				while ($true) {
					#
					if (($verbose -or $debug) -ne $true) { Clear-Host }
					#
					[System.Collections.IEnumerator] $private:cha = $(. script:getOptMap -opts ($private:mea).root -max 1).get_Values().GetEnumerator()
					#
					if (($verbose -or $debug) -ne $true) { Clear-Host }
					#
					if ($private:cha.moveNext() -eq $false) {
						#
						break ;
						#
					}
					#
					$null = $(. script:msxml -xsl $script:intxsl -xml $local:data -out $script:tmpInterXml -param @{"SAP-DR-Recycling-Preprocess.MasterData"="$script:tmpMasterXml"; "SAP-DR-Recycling-Preprocess.Country"="$($private:cha.get_Current())"})
					#
					if([System.IO.Directory]::Exists("$copies")) {
						#
						Copy-Item -Force -Path $script:tmpInterXml -Destination "$copies\inter.xml"
						#
					}
					#
					if (($verbose -or $debug) -ne $true) { Clear-Host }
					#
					[String] $local:res = $(. script:msxml -xsl $script:conxsl -xml $script:tmpInterXml -param @{"SAP-DR-Recycling-Converter.MasterData"="$script:tmpMasterXml"; "SAP-DR-Recycling-Converter.Country"="$($private:cha.get_Current())"})
					#
					Write-Host $local:res
					#
					Read-Host -Prompt "Hit ENTER to continue ..."
					#
				}
				#
			}
			#
		}
		#
	}
	#
	local:cleanup -loc $script:tmpMasterXml
	#
	local:cleanup -loc $script:tmpInterXml
	#
}
#
local:cleanup -loc $script:tmpProductsXml
#
local:cleanup -loc $script:tmpDutiesXml
#
local:cleanup -loc $script:tmpTypesXml
#
local:cleanup -loc $script:tmpFieldsXml
#
$act = "-- void --"
#
if (($verbose -or $debug) -ne $true) { Clear-Host }
#
Write-Verbose -Message @"

SAP-DR-Reporting::main(...): Bye bye ...

"@
# -----------------------------------------------------------------------------------------------
#
#	Werte vor Start des Skripts wiederherstellen
#
$VerbosePreference = $script:saveVerbosePref
$DebugPreference = $script:saveDebugPref
#
# -----------------------------------------------------------------------------------------------
#
