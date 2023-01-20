<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2022)
		
	Revision:
		2021-06-20:	Created to replace 'SAP-DR-Country-List.xslt' for extracting country list from intermediate file in 2-step consolidation.
		2022-10-02:	Major revision tag to reflect readiness for products in lots.
		
-->
<xsl:stylesheet version="1.0" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	exclude-result-prefixes="com dty prd">
	<!--
		
	-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" />
	<!--
		
	-->
	<xsl:key name="cc" use="com:country" match="com:item" />
	<!--
		
	-->
	<xsl:param name="SAP-DR-Country-List.MasterData" select="''" />
	<!--

	-->
	<xsl:variable name="master" select="document ($SAP-DR-Country-List.MasterData)/prd:root" />
	<!--
		
		**********************************************************************
		
	-->
	<xsl:template match="com:report">
		<xsl:variable name="tmp" select="com:item [not(com:country = $master/dty:references/dty:entry/com:countries/com:element)][generate-id (.) = generate-id (key('cc', com:country)[1])]" />
		<xsl:for-each select="$tmp">
			<xsl:message>
				<xsl:text>Overshoot: </xsl:text><xsl:text>	</xsl:text><xsl:value-of select="com:country" />
			</xsl:message>
		</xsl:for-each>
		<root>
			<options default="1">
				<prompt>Für welches Land möchten Sie die Mengen konsolidieren?</prompt>
				<xsl:for-each select="com:item [generate-id (.) = generate-id (key('cc', com:country)[1])][com:country = $master/dty:references/dty:entry/com:countries/com:element]/com:country">
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
</xsl:stylesheet>
