<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"  
	exclude-result-prefixes="atom data meta">
	<!--
		
	-->
	<!--xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" /-->
	<xsl:output method="text" encoding="UTF-8" />
	<!--
		
	-->
	<xsl:template match="meta:properties [data:Group = 'Benutzerdefinierte Inhaltstypen']">
		<xsl:value-of select="data:Id/data:StringValue" /><xsl:text>;</xsl:text><xsl:value-of select="data:Name" /><xsl:text>
</xsl:text>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="/atom:feed">
		<xsl:message><xsl:text>Hallo!</xsl:text></xsl:message>
		<xsl:apply-templates select="atom:entry/atom:content/meta:properties">
			<xsl:sort select="data:Id/data:StringValue" />
		</xsl:apply-templates>
	</xsl:template>
</xsl:stylesheet>
