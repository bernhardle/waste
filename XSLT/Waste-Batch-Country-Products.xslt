<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2024)
		
	Revision:
		2024-03-22:	Created.
		
	Function:
		
		
	Usage:
		XML Formulare\Abfallwirtschaft\ps1\Waste-Reporting-2025.ps1 (copy)
-->
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:xs="http://www.w3.org/2001/XMLSchema" 
	xmlns:fn="http://www.w3.org/2005/xpath-functions" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons"
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty fn prd xs">
	<!--

	-->
	<xsl:param name="Global.country" select="'BE'" />
	<!--

	-->
	<xsl:output method="text" encoding="UTF-8" />
	<!--

	-->
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<!--

	-->
	<xsl:key name="duties" match="dty:references/dty:entry" use="@SPKey"/>
	<xsl:key name="munch" match="prd:product/prd:duties/prd:duty" use="concat (../../@SPKey, '-', @SPKey)" />
	<xsl:key name="countries" match="dty:references/dty:entry/com:countries/com:element" use="." />
	<!--

	-->
	<xsl:template match="prd:product" mode="SCIP">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat ($this/prd:material, ';', key ('duties', @SPKey)[1]/dty:code, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count ($current-group), '#.##0'),'&#xA;')"/>
		</xsl:for-each>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="prd:product" mode="TVVV">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat ($this/prd:material, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count ($current-group), '#.##0'), ';', format-number (sum ($current-group/prd:data/VV_x002d_Alu), '#.##0'), ';', format-number (sum ($current-group/prd:data/VV_x002d_Steel) + sum ($current-group/prd:data/VV_x002d_Tinplate), '#.##0'), ';', format-number (sum ($current-group/prd:data/Paper), '#.##0'), ';', format-number (sum ($current-group/prd:data/VV_x002d_Plastic), '#.##0'), '&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<!--

	-->
	<xsl:template match="prd:product" mode="BATT-WEEE">
		<xsl:param name="batch" />
		<xsl:variable name="this" select="."/>
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$batch = key ('duties', @SPKey)[1]/dty:batch]"/>
			<xsl:value-of select="concat ($this/prd:material, ';', key ('duties', @SPKey)[1]/dty:label, ';', format-number (count ($current-group), '#.##0'), ';', format-number(sum ($current-group/prd:data/Weight), '#.##0,000'), '&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<!--

	-->
	<xsl:template match="/prd:root">
		<xsl:variable name="batch" select="@batch" />
		<list>
			<xsl:choose>
				<xsl:when test="$batch = 'BATT' or $batch = 'WEEE'">
					<xsl:text>material; category; number of parts; total weight of parts [kg]&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $batch]/@SPKey]" mode="BATT-WEEE">
						<xsl:with-param name="batch" select="$batch" />
						<xsl:sort data-type="text" select="prd:material" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:when test="$batch = 'TVVV'">
					<xsl:text>material; category; number of parts; aluminum [gr]; steel/tinplate (FE 04) [gr]; paper [gr]; plastic [gr]&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $batch]/@SPKey]" mode="TVVV">
						<xsl:with-param name="batch" select="$batch" />
					</xsl:apply-templates>
				</xsl:when>
				<xsl:when test="$batch = 'SCIP'">
					<xsl:text>material; category; number of parts&#xA;</xsl:text>
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $batch]/@SPKey]" mode="SCIP">
						<xsl:with-param name="batch" select="$batch" />
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
	<!--

	-->
</xsl:stylesheet>
