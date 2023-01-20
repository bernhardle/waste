<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2021-2022)
		
	Revision:
		2021-02-21:	Created as 'SAP-DR-Country-List.xslt'.
		2022-06-20:	Discontinued.
		2022-12-22:	Revived, extended for csv based XML and renamed to 'SAP-DR-Recycling-Countries.xslt'.
		
-->
<xsl:stylesheet version="1.0" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	xmlns:xls="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:p="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:x="urn:schemas-microsoft-com:office:excel" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	exclude-result-prefixes="com dty p ss x xls">
	<!--
		
	-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--
		
	-->
	<xsl:key name="ccx" use="xls:Cell[7]/xls:Data" match="xls:Row" />
	<xsl:key name="ccc" use="Property [@Name='Empfangsland']" match="Object" />
	<!--
		
	-->
	<xsl:param name="SAP-DR-Recycling-Countries.MasterData" select="''" />
	<!--

	-->
	<xsl:variable name="master" select="document ($SAP-DR-Recycling-Countries.MasterData)/p:root" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="blockWidth" select="90" />
	<!--
		
	-->
	<xsl:template match="*">
		<xsl:message>
			<xsl:value-of select="concat (substring ($constHyp140, 1, $blockWidth), '&#xA;')"/>
			<xsl:text>
 *** ERROR: The input file cannot be processed due to mismatch in data layout. ***
                
</xsl:text>
			<xsl:value-of select="concat (substring ($constHyp140, 1, $blockWidth), '&#xA;')"/>
		</xsl:message>
	</xsl:template>
	<!--
		
		**********************************************************************
		
	-->
	<xsl:template match="xls:Table[xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']">
		<xsl:variable name="tmp" select="xls:Row[position() > 1][not(xls:Cell[7]/xls:Data = $master/dty:references/dty:entry/com:countries/com:element)][generate-id (.) = generate-id (key('ccx', xls:Cell[7]/xls:Data)[1])]" />
		<xsl:for-each select="$tmp">
			<xsl:message>
				<xsl:text>Overshoot: </xsl:text><xsl:text>	</xsl:text><xsl:value-of select="xls:Cell[7]/xls:Data" />
			</xsl:message>
		</xsl:for-each>
		<root>
			<options default="1">
				<prompt>For which country do you want to consolidate the quantities?</prompt>
				<xsl:for-each select="xls:Row[position () > 1][generate-id (.) = generate-id (key('ccx', xls:Cell[7]/xls:Data)[1])][xls:Cell[7]/xls:Data = $master/dty:references/dty:entry/com:countries/com:element]/xls:Cell[7]/xls:Data">
					<xsl:sort select="." />
					<option id="{position()}">
						<key><xsl:value-of select="." /></key>
						<label><xsl:value-of select="." /></label>
					</option>
				</xsl:for-each>
			</options>
		</root>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Table" />
	</xsl:template>
	<!--
			
			Verteiler Template für Excel 2003 xml Format
			
	-->
	<xsl:template match="/xls:Workbook[xls:Worksheet/@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']/xls:Table" />
	</xsl:template>
	<!--
		
			Berechnungstemplate für XML aus csv Datei
		
	-->
	<xsl:template match="/Objects | Objects">
		<xsl:variable name="tmp" select="Object [not(Property [@Name='Empfangsland'] = $master/dty:references/dty:entry/com:countries/com:element)][generate-id (.) = generate-id (key('ccc', Property [@Name='Empfangsland'])[1])]" />
		<xsl:for-each select="$tmp">
			<xsl:message>
				<xsl:text>Overshoot: </xsl:text><xsl:text>	</xsl:text><xsl:value-of select="Property [@Name='Empfangsland']" />
			</xsl:message>
		</xsl:for-each>
				<root>
			<options default="1">
				<prompt>For which country do you want to consolidate the quantities?</prompt>
				<xsl:for-each select="Object [generate-id (.) = generate-id (key('ccc', Property [@Name='Empfangsland'])[1])][Property [@Name='Empfangsland'] = $master/dty:references/dty:entry/com:countries/com:element]/Property [@Name='Empfangsland']">
					<xsl:sort select="." />
					<option id="{position()}">
						<key><xsl:value-of select="." /></key>
						<label><xsl:value-of select="." /></label>
					</option>
				</xsl:for-each>
			</options>
		</root>
	</xsl:template>
	<!--

	-->
	<xsl:template match="/Wrapped">
		<xsl:apply-templates select="Objects" />
	</xsl:template>
</xsl:stylesheet>
