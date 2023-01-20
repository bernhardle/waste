#
# REST Abfrage https://rowz.sharepoint.com/teams/UFI2021-DEEARMengenmeldungenGerteundLampen/_api/web/lists(guid'bbc5a316-a639-4031-9b4e-1c4d293b3eaa')/items?$select=Title,Gewicht,Kategorie,Legacy,Modified&$top=10000&$orderby=ID
#
param([String] $sda="$pwd\EAR-Stammdaten.xml", [Boolean] $verbose = $false, [Boolean] $debug = $false)
#
# -----------------------------------------------------------------------------------------------
#
[String] $script:stylesheet = @"
<?xml version="1.0" encoding="UTF-8"?>
<!--

-->
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
	xmlns:ear="http://www.stiftung-elektroaltgeraete-register.de"
	xmlns:dsd="http://www.gruener-punkt.de"
	xmlns:xlink="http://www.w3.org/1999/xlink"
	exclude-result-prefixes="xls">

	<xsl:param name="debug" select="false()" />
    <xsl:param name="EAR-DSD-Converter.Country" select="'DE'" />
	<xsl:param name="EAR-DSD-Converter.Mode" select="'void'" />
	<xsl:param name="EAR-DSD-Converter.DSD-Date" select="'01.01.1970'" />
	<xsl:param name="EAR-DSD-Converter.EAR-Stammdaten" />
	
	<xsl:variable name="constBlank140" select="'                                                                                                                                            '" />
	<xsl:variable name="constHyphen140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="constDot140" select="'............................................................................................................................................'" />
	<xsl:variable name="blockWidth" select="90" />
	
	<xsl:variable name="ear-stammdaten" select="document (`$EAR-DSD-Converter.EAR-Stammdaten)/ns1:feed" />
	<xsl:variable name="base" select="`$ear-stammdaten/ns1:entry" />
	
	<!-- xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/ -->	

	<xsl:key name="b2c" use="ns1:content/m:properties/d:Kategorie" match="ns1:entry" />
	<xsl:key name="mat" use="xls:Cell[1]/xls:Data" match="xls:Row" />

	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	
	<xsl:output method="text" encoding="asci" />
	<!--
		Klartextbeschreibung zu den RO Kategorien 0...6 und wie sich diese auf
		die von der EAR verwendeten Kategorien abbilden. Die Texte werden
		nach den im Attribut genannten 'key' zeilenweise am Anfang der Gruppe
		ausgegeben.
	-->
	<xsl:template name="EAR-Kategorie-Beschreibung">
		<xsl:param name="gruppe" />
		<xsl:param name="indent" select="21" />
		<xsl:choose>
			<xsl:when test="number (`$gruppe) = number(0)">
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 0', substring (`$constBlank140, 1, 6), '&quot;Fehlerkategorie sammelt Produkte mit invaliden&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'Stammdaten. Einzelheiten dazu sind im Abschnitt&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'zum Stammdaten-Check.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number (1)">
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 1', substring (`$constBlank140, 1, 6), '&quot;EAR Kategorie 1: Waermeuebertraeger, Heizungen.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number (2)">
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 2', substring (`$constBlank140, 1, 6), '&quot;EAR Kategorie  2.1: Bildschirme, Monitore und&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'Geraete, die Bildschirme mit einer Oberflaeche&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'von mehr als 100 cm2 enthalten.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number(3)">
					<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 3', substring (`$constBlank140, 1, 6), 'EAR Kategorie  3.2: Lampen, die in privaten&#xA;')" />
					<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'Haushalten genutzt werden koennen ausser&#xA;')" />
					<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'Gasentladungslampen.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number (4)">
					<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 4', substring (`$constBlank140, 1, 6), '&quot;EAR Kategorie 4.1: Grossgeraete, bei denen mind.&#xA;')" />
					<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'eine der auesseren Abmessungen mehr als 50cm&#xA;')" />
					<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'betraegt, ausg. grosse Photovoltaikmodule.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number (5)">
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 5', substring (`$constBlank140, 1, 6), '&quot;EAR Kategorie 5.1: Kleingeraete, ausgenommen&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'kleine Photovoltaikmodule.&quot;&#xA;')" />
			</xsl:when>
			<xsl:when test="number (`$gruppe) = number (6)">
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 5), 'Gruppe: 6', substring (`$constBlank140, 1, 6), '&quot;EAR Kategorie 6.1: Kleine Geraete der Informations-&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'und Telekommunikationstechnik bei denen keine der&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, `$indent), 'auesseren Abmessungen mehr als 50cm betraegt.&quot;&#xA;')" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:message>
					<xsl:text>[FATAL] Invalide Eingruppierung des Materials.</xsl:text>
				</xsl:message>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:value-of select="'&#xA;'" />
	</xsl:template>

	<xsl:template name="recSumWeights">
		<xsl:param name="items" />
		<xsl:param name="sum" select="0" />
		<xsl:choose>
			<xsl:when test="`$items">
				<xsl:variable name="pos" select="`$base [`$items[1]/xls:Cell[1]/xls:Data = ns1:content/m:properties/d:Title]" />
				<xsl:variable name="tmp">
					<xsl:choose>
						<xsl:when test="not (`$pos/ns1:content/m:properties/d:Gewicht)">
							<xsl:value-of select="`$sum" />
						</xsl:when>
						<xsl:when test="not (number (`$pos/ns1:content/m:properties/d:Gewicht) = number (`$pos/ns1:content/m:properties/d:Gewicht))">
							<xsl:value-of select="`$sum" />
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="`$sum + `$items [1]/xls:Cell[5]/xls:Data * `$pos/ns1:content/m:properties/d:Gewicht" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:call-template name="recSumWeights">
					<xsl:with-param name="items" select="`$items [position() > 1]" />
					<xsl:with-param name="sum" select="`$tmp" />
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="`$sum" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="*">
		<xsl:message>
			<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
			<xsl:text>
 *** FEHLER: Die XML Quelldatei kann nicht transformiert werden. ***
 
      Die Arbeitsmappe hat keine Tabelle mit dem Namen &apos;Sheet1&apos; in
      der die erste Zeile in der ersten Zelle den Wert &apos;Material&apos; und
      in der fuenften Zelle den Wert &apos;Fakturierte Menge&apos; enthaelt.
                
</xsl:text>
			<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
		</xsl:message>
	</xsl:template>
	
	<xsl:template match="*" mode="run">
		<xsl:apply-templates select="." />
	</xsl:template>
	
	<xsl:template match="*" mode="check1">
		<xsl:apply-templates select="." />
	</xsl:template>
	
	<xsl:template match="*" mode="check2">
		<xsl:apply-templates select="." />
	</xsl:template>
	
	<xsl:template match="xls:Row">
		<xsl:variable name="total" select="sum (key('mat', xls:Cell[1]/xls:Data)[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country]/xls:Cell[5])" />
		<xsl:variable name="pos" select="`$base [current ()/xls:Cell[1]/xls:Data = ns1:content/m:properties/d:Title]" />
		<!--

		-->
		<xsl:variable name="col1w" select="6" />
		<xsl:variable name="col2w" select="12" />
		<xsl:variable name="col3w" select="46" />
		<xsl:variable name="col4w" select="8" />
		<xsl:variable name="col5w" select="6" />
		<xsl:variable name="col6w" select="8" />
		<!--

		-->
		<xsl:value-of select="substring (`$constBlank140, 1, `$col1w)" />
		<xsl:value-of select="xls:Cell[1]/xls:Data" />
		<xsl:value-of select="substring (`$constBlank140, 1, `$col2w - string-length (xls:Cell[1]/xls:Data))" />
		<xsl:value-of select="substring (xls:Cell[2]/xls:Data, 1, `$col3w)" />
		<xsl:choose>
			<xsl:when test="`$pos">
				<xsl:value-of select="substring (`$constBlank140, 1, `$col4w + `$col3w - string-length (substring (xls:Cell[2]/xls:Data, 1, `$col3w)) - string-length (`$pos/ns1:content/m:properties/d:Gewicht) - 3)" />
				<xsl:value-of select="concat (translate(`$pos/ns1:content/m:properties/d:Gewicht, '.,',',.'), ' kg')" />
				<xsl:value-of select="substring (`$constBlank140, 1, `$col5w - string-length (`$pos/ns1:content/m:properties/d:Kategorie))" />
				<xsl:value-of select="`$pos/ns1:content/m:properties/d:Kategorie" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="substring (`$constBlank140, 1, `$col5w + `$col4w + `$col3w - string-length (substring (xls:Cell[2]/xls:Data, 1, `$col3w)))" />
			</xsl:otherwise>
		</xsl:choose>
		<xsl:value-of select="substring (`$constBlank140, 1, `$col6w - string-length (format-number (`$total, '#.##0')))" />
		<xsl:value-of select="format-number (`$total, '#.##0')" />
		<xsl:text>
</xsl:text>	
	</xsl:template>
	
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']" mode="check1">
		<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
		<xsl:variable name="nobase" select="xls:Row[position () > 1][xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][generate-id (key('mat', xls:Cell[1]/xls:Data)[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][1]) = generate-id (.)][not(xls:Cell[1]/xls:Data = `$base/ns1:content/m:properties/d:Title)]" />
		<xsl:for-each select="`$nobase">
			<xsl:if test="position () = 1">
				<xsl:text>
 *** Produkte in Belegen, zu denen keine WEEE-Stammdaten angelegt sind:

</xsl:text>
			</xsl:if>
			<xsl:apply-templates select="." />
		</xsl:for-each>
		<xsl:text>
</xsl:text>
	</xsl:template>
	<!--
		Stammdaten Check:
	-->
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']" mode="check2">
		<!--
		-->
		<xsl:variable name="malbase" select="`$base [not (ns1:content/m:properties/d:Gewicht > 0 and (translate(ns1:content/m:properties/d:Kategorie, '123456d','ddddddx') = 'd'))]" />
		<xsl:for-each select="xls:Row[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][generate-id (key('mat', xls:Cell[1]/xls:Data)[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][1]) = generate-id (.)][xls:Cell[1]/xls:Data = `$malbase/ns1:content/m:properties/d:Title]">
			<xsl:if test="position () = 1">
				<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
				<xsl:text>
 *** Produkte in Belegen, zu denen die WEEE-Stammdaten unvollstaendig sind:

</xsl:text>
			</xsl:if>
			<xsl:apply-templates select="." />
		</xsl:for-each>
		<xsl:text>
</xsl:text>
	</xsl:template>
	<!--
		Legacy Check
	-->
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']" mode="check3">
		<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
        <xsl:variable name="sdalcy" select="`$base/ns1:content/m:properties[d:Legacy = 'true']/d:Title" />
		<xsl:variable name="legacy" select="xls:Row[position () > 1][xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][generate-id (key('mat', xls:Cell[1]/xls:Data)[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][1]) = generate-id (.)][xls:Cell[1]/xls:Data = `$sdalcy]" />
		<xsl:for-each select="`$legacy">
			<xsl:if test="position () = 1">
				<xsl:text>
 *** Produkte in Belegen, zu denen Legacy WEEE-Stammdaten angezogen werden:

</xsl:text>
			</xsl:if>
			<xsl:apply-templates select="." />
		</xsl:for-each>
		<xsl:text>
</xsl:text>
	</xsl:template>
	<!--
		Berechnung der verkauften Gesamtmengen
	-->
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']" mode="ear-run">
		<xsl:variable name="table" select="." />
		<xsl:value-of select="concat (substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
		<xsl:text>
 *** Verkaufte Elektrogeraete nach Kategorien in Stueckzahlen und Gesamtgewicht:
 
</xsl:text>
		<xsl:for-each select="`$base[generate-id (key ('b2c', ns1:content/m:properties/d:Kategorie)[1]) = generate-id (.)]">
			<xsl:sort data-type="number" select="ns1:content/m:properties/d:Kategorie"  order="ascending" />
			<xsl:variable name="catprods" select="key ('b2c', ns1:content/m:properties/d:Kategorie)" />
			<xsl:variable name="items" select="`$table/xls:Row[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][xls:Cell[1]/xls:Data = `$catprods/ns1:content/m:properties/d:Title]" />
			<!--
				
			-->
			<xsl:if test="`$items">
				<!--
					
				-->
				<xsl:variable name="total" select="sum (`$items/xls:Cell[5]/xls:Data)" />
				<xsl:variable name="weight">
					<xsl:call-template name="recSumWeights">
						<xsl:with-param name="items" select="`$items" />
					</xsl:call-template>
				</xsl:variable>
				<!--

				-->
				<xsl:variable name="indent" select="45" />
				<xsl:call-template name="EAR-Kategorie-Beschreibung">
					<xsl:with-param name="gruppe" select="`$catprods[1]/ns1:content/m:properties/d:Kategorie" />
					<xsl:with-param name="indent" select="21" />
				</xsl:call-template>
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 21), 'Anzahl Belege: ', substring (`$constBlank140, 1, `$indent - 14 - string-length (format-number (count (`$items), '#.##0'))), format-number (count (`$items), '#.##0'), '&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 21), 'Anzahl verkaufte Produkte: ', substring (`$constBlank140, 1, `$indent - 26 - string-length (format-number (`$total, '#.##0'))), format-number (`$total, '#.##0'), '&#xA;')" />
				<xsl:value-of select="concat (substring (`$constBlank140, 1, 21), 'Gesamtgewicht: ', substring (`$constDot140, 1, `$indent - 15 - string-length (format-number(`$weight, '#.##0'))), ' ', format-number(`$weight, '#.##0,00'), ' kg', '&#xA;')" />
				<!--

				-->
				<xsl:text>
</xsl:text>
			</xsl:if>
		</xsl:for-each>
		<xsl:value-of select="substring (`$constHyphen140, 1, `$blockWidth)" />
		<xsl:text>
</xsl:text>
	</xsl:template>
	
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']" mode="dsd-run">
		<xsl:variable name="table" select="." />
		<xsl:message>
			<xsl:value-of select="concat ('&#xA;', substring (`$constHyphen140, 1, `$blockWidth), '&#xA;', '&#xA;')" />
			<xsl:text> *** Verpackungen im Zeitraum konsolidiert nach Materialnummer in Stueck:</xsl:text>
			<xsl:value-of select="concat ('&#xA;', '&#xA;', substring (`$constHyphen140, 1, `$blockWidth), '&#xA;')" />
		</xsl:message>
		<xsl:variable name="semic" select="';'" />
		<xsl:value-of select="concat('EAN', `$semic, 'VERPACKUNGSNUMMER', `$semic, 'ZUORDNUNGSKRITERIEN', `$semic, 'PERIODE', `$semic, 'STUECK', '&#xD;&#xA;')" />
		<xsl:for-each select="xls:Row[position () > 1][xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][generate-id (key('mat', xls:Cell[1]/xls:Data)[xls:Cell[7]/xls:Data = `$EAR-DSD-Converter.Country][1]) = generate-id (.)]">
			<xsl:value-of select="concat (`$semic, xls:Cell[1]/xls:Data, `$semic, `$semic, `$EAR-DSD-Converter.DSD-Date, `$semic, sum (key('mat', xls:Cell[1]/xls:Data)/xls:Cell[5]/xls:Data), '&#xD;&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<!--
		Verteiler Template
	-->
	<xsl:template match="/xls:Workbook[xls:Worksheet/@ss:Name = 'Sheet1']">
		<xsl:choose>
			<xsl:when test="`$EAR-DSD-Converter.Mode = 'ear'">
				<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']" mode="ear" />
			</xsl:when>
			<xsl:when test="`$EAR-DSD-Converter.Mode = 'dsd'">
				<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']" mode="dsd" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="concat ('&#xA;', substring (`$constHyphen140, 1, `$blockWidth), '&#xA;', '&#xA;')" />
				<xsl:value-of select="concat (' *** FEHLER: Der parametrisierte Modus &quot;', `$EAR-DSD-Converter.Mode, '&quot; ist nicht vorgesehen. ***', '&#xA;')" />
				<xsl:value-of select="concat ('&#xA;', substring (`$constHyphen140, 1, `$blockWidth), '&#xA;', '&#xA;')" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']" mode="ear">
		<xsl:apply-templates select="xls:Table" mode="check1" />
		<xsl:apply-templates select="xls:Table" mode="check2" />
		<xsl:apply-templates select="xls:Table" mode="check3" />
		<xsl:apply-templates select="xls:Table" mode="ear-run" />
	</xsl:template>
	
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']" mode="dsd">
		<xsl:apply-templates select="xls:Table" mode="dsd-run" />
	</xsl:template>
	
</xsl:stylesheet>
"@
#
# -----------------------------------------------------------------------------------------------
#
function script:delete([String] $del = $(throw "[Fatal] SAP-EAR-convert.ps1::delete() Filename missing."))
{
	if ([System.IO.File]::Exists($del) -eq $true) {
		Write-Debug @"
[Debug] SAP-EAR-DSD-Converter.ps1::delete() Deleting file '$del'.
"@
		[System.IO.File]::Delete($del)
	}
}
#
[Reflection.Assembly] $local:ass = [Reflection.Assembly]::LoadWithPartialName("System.Xml")
#
# -----------------------------------------------------------------------------------------------
#
function script:getOptMap([System.Xml.XmlElement] $opts, [Int] $max = 1, [String] $brk = "x")
{
	#
	[String] $local:title = @"
Was möchten Sie tun?
"@
	[System.Collections.Hashtable] $private:hsh = @{}
	#
	if($opts.options.prompt -ne $null) {
		$title = $opts.options.prompt
	}
	#
	[System.Text.StringBuilder] $local:msg = new-object -typeName System.Text.StringBuilder -argumentList @"

[Interactive]
	
	$title

"@
	#
	foreach($local:pos in $opts.options.option) {
		$hsh[[Int]$pos.id] = ($pos.key, $pos.label)
		$null = $msg.append(@"
	
	[$($pos.id)]	$($hsh[[Int]$pos.id][1])
"@)
	}
	#
	$null = $msg.append(@"
	
	
	[$brk]	Abbrechen

"@)
	[String] $trailer = @"

Bitte die zu bearbeitende(n) Nummer(n) [#] ggf. mit Komma getrennt eingeben (max. $max)
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
			if ($debug -ne $true) {
			#	$null = Clear-Host
			}
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
				if ($item -as [Int] -ne $null) {
					[Int] $local:id = [Int] $item
					if($hsh.contains($id)) {
						if($res.contains($id) -eq $false) {
							$res[$id] = $hsh[$id][0]
						}
					} else {
						$lasterr = @"
$lasterr
						
	***	FEHLER: Die Nummer $id ist in der Liste nicht enthalten.	***

"@
						$end = $false
					}
				} else {
					$lasterr = @"
$lasterr
					
	***	FEHLER: Die Auswahl $item ist keine Zahl.	***

"@
					$end = $false
				}
			}
		}
		#
		if ($res.get_Count() -gt $max) {
			$lasterr = @"
$lasterr
			
	***	FEHLER: Die Auswahl enthält zu viele Einträge (max. $max).	***

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
function script:OpenFileDialog([String] $title = "Oups.", [String] $type = "all", [String] $defpath = "")
{
	#
	[Reflection.Assembly] $ass = [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	#
	[System.Windows.Forms.OpenFileDialog] $openFileDialog1 = new-object -typeName System.Windows.Forms.OpenFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'xml'='XML Datei (*.xml)|*.xml'; 'htm'='HTML Datei (*.htm, *.html)|*.htm;*.html'; 'all'='Alle Dateien (*.*)|*.*'}
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
function script:SaveFileDialog([String] $title = "Datei speichern", [String] $type = 'csv', [String] $defpath = "", [String] $defname = $null, [Boolean] $delete = $false)
{
    #
	[Reflection.Assembly] $local:ass = [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	#
	[System.Windows.Forms.SaveFileDialog] $local:saveFileDialog1 = new-object -typeName System.Windows.Forms.SaveFileDialog
	#
	[System.Collections.Hashtable] $local:filters = @{'csv'='CSV Datei (*.csv)|*.csv' ; 'xml'='XML Datei (*.xml)|*.xml'; 'xsl' = 'XSLT Stylesheet (*.xsl, *.xslt)|*.xsl,*.xslt'; 'htm'='HTML Datei (*.htm, *.html)|*.htm;*.html'; 'pptx'='Powerpoint Slideshow (*.pptx)|*.pptx'; 'pdf'='Portable Data Format (*.pdf)|*.pdf'; 'fo'='Formatting Objects (*.fo)|*.fo'; 'all'='Alle Dateien (*.*)|*.*'}
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
	if($defname -ne $null) {
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
function script:msxml([String] $xml = $(throw "[Fatal] msxml.ps1: Filename for XML input missing."), [System.Collections.Hashtable] $param = @{}) {
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
	if($param -ne $null) {
		[String[]] $local:val = $null
		foreach($val in $param.get_Keys()) {
			Write-Debug @"
[DEBUG] SAP-EAR-DSD-Converter::msxml(...): Adding XSLT parameter $val = $($param[$val])
"@
			$xal.AddParam($val, "", $($param[$val]))
		}
	}
	#
	$xslt.Load([System.Xml.XmlReader]::Create($(New-Object -typeName System.IO.StringReader -argumentList $stylesheet)), $xset, $xres)
	# $xslt.Load($xsl, $xset, $xres)
	#
	[System.Xml.XPath.XPathDocument] $local:xpd = New-Object -typename System.Xml.XPath.XPathDocument -argumentlist $xml
	[System.IO.StringWriter] $local:osw = New-Object System.IO.StringWriter
	$xslt.Transform($xpd, $xal, $osw)
	$osw.close()
	[String] $local:res = $osw.ToString()
	$osw.Dispose()
	#
	return $res
}
# -----------------------------------------------------------------------------------------------
#
#
#	Werte der Variablen $VerbosePreference und $DebugPreference sichern
#	und entsprechend den Parametern $verbose und $debug neu setzen.
#
[System.Management.Automation.ActionPreference] $script:saveVerbosePref = $VerbosePreference
[System.Management.Automation.ActionPreference] $script:saveDebugPref = $DebugPreference
#
if($verbose -eq $true -or  $debug -eq $true) {
	Write-Verbose "[INFO] SAP-EAR-convert.ps1::main(): verbose mode."
	$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
    if($debug -eq $true) {
	    Write-Verbose @"
[INFO] SAP-EAR-DSD-Converter.ps1::main(): debug mode.
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
#
[Xml] $local:table = @"
<?xml version="1.0" encoding="UTF-8"?>
<root>
	<options default="1">
		<prompt>Was möchten Sie aus dem SAP Export (Excel 2003 XML) machen?</prompt>
		<option id="1">
			<key>1</key>
			<label>EAR: In konsolidierte Gewichte für EAR Portal konvertieren ...</label>
		</option>
		<option id="2">
			<key>2</key>
			<label>DSD: In Datei (CSV Format) für DSD Import konvertieren ...</label>
		</option>
		<option id="3">
			<key>3</key>
			<label>DEBUG: Das Stylesheet ausspucken ...</label>
		</option>
	</options>
</root>
"@
#
#
[Xml] $local:tableCC = @"
<?xml version="1.0" encoding="UTF-8"?>
<root>
	<options default="1">
		<prompt>Für welches Land möchten Sie die Konvertierung ausführen?</prompt>
		<option id="1">
			<key>DE</key>
			<label>Deutschland</label>
		</option>
		<option id="2">
			<key>FR</key>
			<label>Frankreich</label>
		</option>
	</options>
</root>
"@
#
#
while($true) {
	#
	Clear-Host
	#
	[System.Collections.IEnumerator] $local:choice = $(. getOptMap -opts ($local:table).root -max 1).get_Values().GetEnumerator()
	#
	if ($local:choice.moveNext() -eq $false) {
		#
		break ;
		#
	}
	#
	[String] $local:load = $(. OpenFileDialog -title "SAP Excel 2003 Datei auswählen" -type "xml" -defpath $pwd -verbose $verbose -debug $debug)
	#
	if ([System.IO.File]::Exists($local:load) -eq $true) {
		#
        # Das Verzeichnis der gewählten Datei als
        # Startpunkt für die Suche nach dem Ablageort
        # der konvertierten Datei und für die nächste
        # Runde festlegen.
        #
        $pwd = $(Get-Item $local:load).DirectoryName
        #
		switch([Int]$local:choice.get_Current()) {
			#
			1	{ 
				#
				[String] $local:text = $(. msxml -xsl $script:xsl -xml $local:load -debug $debug -verbose $verbose -param @{"EAR-DSD-Converter.EAR-Stammdaten"="$sda"; "EAR-DSD-Converter.Country"="DE"; "EAR-DSD-Converter.Mode"="ear"})
				#
				Write-Host $local:text
				#
			}
			2	{
				#
				#	Default Berichtsperiode ist der 1. des Vormonats, weil die Meldung immer im Folgemonat gemacht wird.
				#
				[DateTime] $private:periode = [DateTime]::Today.AddMonths(-1).AddDays(1 - [DateTime]::Today.Day)
				#
				do {
					[String] $private:date = Read-Host -Prompt @"
				
        Geben Sie das Datum der Periode der Importdaten ein oder übernehmen
        Sie das voreingestellte Datum '$("{0:d}" -f $periode)' mit der Eingabetaste
"@
					#
					if ($private:date.length -eq 0) {
						break ;
					}
					#
				} while (![DateTime]::TryParse($private:date, [ref] $private:periode))
				#
				[String] $private:text = $(. msxml -xml $local:load -debug $debug -verbose $verbose -param @{"EAR-DSD-Converter.Mode"="dsd";"EAR-DSD-Converter.DSD-Date"="$('{0:d}' -f $periode)"}) 
				#
				ConvertFrom-Csv -Delimiter ';' -InputObject $private:text | Format-Table -AutoSize 
				#
				[String] $private:save = $(. SaveFileDialog -defpath $pwd -defname "$('{0:MM-yyyy}' -f $periode).csv")
				#
				if($private:save -ne $null -and $private:save -ne "") {
					#
					Out-File -InputObject $private:text -Encoding ascii -FilePath $private:save
					#
				}
				#
			}
			#
			3	{
                #
                [String] $private:save = $(. SaveFileDialog -defpath $pwd -type 'xsl' -defname 'Stylesheet.xslt')
                #
                if($private:save -ne $null -and $private:save -ne "") {
                    #
                    $null = $script:stylesheet | Out-File -FilePath $private:save -Encoding 'utf8'
                    #
                }
                #
			}
		}
		#
		Read-Host -Prompt @"
Fertig. Zum Fortsetzen die Eingabetaste drücken. Die Ausgabe wird dann gelöscht
"@
		#
	}
	#
}
# -----------------------------------------------------------------------------------------------
#
#	Werte vor Start des Skripts wiederherstellen
#
$VerbosePreference = $script:saveVerbosePref
$DebugPreference = $script:saveDebugPref
#
# -----------------------------------------------------------------------------------------------
#
Write-Host @"

Bye bye!

"@
#