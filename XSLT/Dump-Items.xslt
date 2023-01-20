<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:a="http://www.w3.org/2005/Atom" 
	xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml">
	<!--
		
	-->
	<xsl:output method="text" encoding="UTF-8" indent="yes"/>
	<!--
		https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/items?$select=Id,Material,Duty_x002d_ListId,Description1&$top=10000&$orderby=ID&$filter=ContentTypeId%20eq%20%270x01003FAF714C6769BF4FA1B36DCF47ED6597030100F1F1D05B8F05D14082CB1F3E0D35D4E3%27
	-->
	<xsl:template match="a:* | m:* | d:* [@m:null='true'] | d:Id | d:GUID | d:FileSystemObjectType | d:ServerRedirectedEmbedUri | d:ServerRedirectedEmbedUrl | d:ContentTypeId | d:ComplianceAssetId" />
	
	<xsl:template match="a:* | m:* | d:* [@m:null='true'] | d:Id | d:GUID | d:FileSystemObjectType | d:ServerRedirectedEmbedUri | d:ServerRedirectedEmbedUrl | d:ContentTypeId | d:ComplianceAssetId" mode="header" />
	
	<xsl:template match="d:element">
		<xsl:text>|</xsl:text>
		<xsl:value-of select="." />
		<xsl:if test="position () = last ()">
			<xsl:text>|</xsl:text>
		</xsl:if>
	</xsl:template>
		
	<xsl:template match="d:* [@m:type='Collection(Edm.Int32)']">
		<xsl:text>&quot;</xsl:text>
		<xsl:apply-templates select="child::node()">
			<xsl:sort select="." order="ascending" data-type="number" />
		</xsl:apply-templates>
		<xsl:text>&quot;</xsl:text>
		<xsl:if test="position () != last ()"><xsl:text>;</xsl:text></xsl:if>
	</xsl:template>

	<xsl:template match="d:*">
		<xsl:text>&quot;</xsl:text><xsl:value-of select="." /><xsl:text>&quot;</xsl:text>
		<xsl:if test="position () != last ()"><xsl:text>;</xsl:text></xsl:if>
	</xsl:template>

	<xsl:template match="m:properties">
		<xsl:apply-templates select="d:*" /><xsl:text>
</xsl:text>
	</xsl:template>
	
		<xsl:template match="m:properties" mode="header">
		<xsl:apply-templates select="d:*" mode="header" /><xsl:text>
</xsl:text>
	</xsl:template>
	
	<xsl:template match="d:*" mode="header">
		<xsl:text>&quot;</xsl:text><xsl:value-of select="local-name(.)" /><xsl:text>&quot;</xsl:text>
		<xsl:if test="position () != last ()"><xsl:text>;</xsl:text></xsl:if>
	</xsl:template>
	
	<xsl:template match="/a:feed">
		<xsl:apply-templates select="a:entry [1]/a:content/m:properties" mode="header" />
		<xsl:apply-templates select="a:entry/a:content/m:properties" />
	</xsl:template>
	
</xsl:stylesheet>
