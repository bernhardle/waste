<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:a="http://www.w3.org/2005/Atom" 
	xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml">

	<xsl:param name="insert-name" select="'ItemListId'" />
	<xsl:param name="insert-value" select="1758" />
	
	<xsl:output method="text" encoding="UTF-8" indent="yes"/>
		
	<xsl:template match="a:* | m:* | d:*" />

	<xsl:template match="d:Part_x002d_ListId [d:element] | d:ItemListId [d:element] | d:Duty_x002d_ListId [d:element]">
		<xsl:text>{&quot;</xsl:text><xsl:value-of select="local-name(.)" /><xsl:text>&quot; : [</xsl:text>
		<xsl:for-each select="d:element">
			<xsl:value-of select="." />
			<xsl:if test="not (position () = last ())">
				<xsl:text>,</xsl:text>
			</xsl:if>
		</xsl:for-each>
		<xsl:if test="not(d:element = $insert-value)">
			<xsl:if test="d:element">
				<xsl:text>,</xsl:text>
			</xsl:if>
			<xsl:value-of select="$insert-value" />
		</xsl:if>
		<xsl:text>]}</xsl:text>
	</xsl:template>

	<xsl:template match="m:properties">
		<xsl:apply-templates select="d:* [local-name() = $insert-name]" />
	</xsl:template>
	
	<xsl:template match="/a:feed">
		<xsl:if test="count(a:entry) > 1">
			<xsl:message terminate="yes">
				<xsl:text>FATAL: Multiple hits</xsl:text>
			</xsl:message>
		</xsl:if>
		<xsl:apply-templates select="a:entry/a:content/m:properties" />
	</xsl:template>
	
</xsl:stylesheet>
