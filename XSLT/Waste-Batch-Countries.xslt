<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2024)
		
	Revision:
		2024-02-05:	Created as 'Waste-Batch-Countries.xslt'.
		
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
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone="yes" />
	<!--

	-->
	<xsl:key name="countries" match="dty:references/dty:entry/com:countries/com:element" use="." />
	<!--

	-->
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
	<!--

	-->
</xsl:stylesheet>
