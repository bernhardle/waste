<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:xs="http://www.w3.org/2001/XMLSchema"
	xmlns:fn="http://www.w3.org/2005/xpath-functions"
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata"
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons"
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty fn prd xs">
	
	<xsl:param name="Global.country" select="'DE'" />
	<xsl:param name="Global.batch" select="'BATT'" />
	
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	
	<xsl:key name="duties" match="dty:references/dty:entry" use="@SPKey" />
	
	<xsl:template match="prd:product">
		<xsl:param name="duties" />
		<xsl:if test="prd:duties/prd:duty[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][key ('duties', @SPKey)[1]/dty:batch = $Global.batch]">
			<product>
				<material>
					<xsl:value-of select="prd:material" />
				</material>
				<kurztext>
					<xsl:value-of select="prd:materialkurztext" />
				</kurztext>
				<xsl:for-each-group select="prd:duties/prd:duty[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][key ('duties', @SPKey)[1]/dty:batch = $Global.batch]" group-by="@SPKey">
					<xsl:element name="{key ('duties', @SPKey)[1]/dty:batch}">
						<category>
							<xsl:value-of select="key ('duties', @SPKey)[1]/dty:label" />
						</category>
						<count>
							<xsl:value-of select="count (current-group ())" />
						</count>
						<weight>
							<xsl:value-of select="sum (current-group ()/prd:data/Weight)" />
						</weight>
					</xsl:element>
				</xsl:for-each-group>
			</product>
		</xsl:if>
	</xsl:template>
	
	<xsl:template match="/prd:root">
		<products>
			<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country]/@SPKey]" />
		</products>
	</xsl:template>
	
</xsl:stylesheet>
