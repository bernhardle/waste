<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2025)
		
	Revision:
		2025-06-23:	Created as 'Waste-Items-Feed-In.xslt'
		
	Purpose:
		
	Usage:
		
-->
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	exclude-result-prefixes="atom data meta">
	<!--

	-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone="yes" />
	<!--

	-->
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="Waste-Items-Feed-In.ProductLoadFile" />
	<xsl:param name="Waste-Items-Feed-In.ContentTypesLoadFile" />
	<xsl:param name="Waste-Items-Feed-In.Product-ContentTypeID" /> <!-- select="'0x01003FAF714C6769BF4FA1B36DCF47ED659702'" / -->
	<!--
		Herunterladen der Dateien mit:
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/ContentTypes
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/Fields?$select=InternalName,Title,Description,TypeAsString,TypeDisplayName,TypeShortDescription
	-->
	<xsl:variable name="Waste-Items-Feed-In.Products" select="document($Waste-Items-Feed-In.ProductLoadFile)/atom:feed/atom:entry" />
	<xsl:variable name="Waste-Items-Feed-In.ContentTypes" select="document($Waste-Items-Feed-In.ContentTypesLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<!--
		entry: indexes the ID to the entry with the ID as primary key
		back: indexes the ID to foreign reference fields (item list, part list, reference item) containing the ID
	-->
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry" />
	<xsl:key name="back" use="atom:content/meta:properties/data:ItemListId/data:element | atom:content/meta:properties/data:Part_x002d_ListId/data:element | atom:content/meta:properties/data:Reference_x002d_ProductId" match="/atom:feed/atom:entry" />
	<xsl:key name="items" use="Key" match="Entry" />
	<!--
		
	-->
	<xsl:template match="meta:properties">
		<xsl:param name="item" />
		<Entry>
			<Key>
				<xsl:value-of select="concat (data:ID, '-', $item/atom:content/meta:properties/data:ID)" />
			</Key>
			<Material>
				<xsl:value-of select="data:Material" />
			</Material>
			<Kurztext>
				<xsl:value-of select="data:Description1" />
			</Kurztext>
			<Item>
				<xsl:value-of select="$item/atom:content/meta:properties/data:Material" />
			</Item>
		</Entry>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="atom:entry">
		<xsl:param name="item" select="." />
		<xsl:param name="loop" select="'[x]'" />
		<xsl:variable name="top" select="key ('back', atom:content/meta:properties/data:Id)" />
		<xsl:choose>
			<xsl:when test="contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] Waste-Items-Feed-In.xslt (line 461): Loop detected in chain </xsl:text><xsl:value-of select="concat('[', atom:content/meta:properties/data:Id, ']-', $loop)" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>

				<xsl:if test="$top">
					<xsl:apply-templates select="$top">
						<xsl:with-param name="item" select="$item" />
						<xsl:with-param name="loop" select="concat('[', atom:content/meta:properties/data:Id, ']-', $loop)" />
					</xsl:apply-templates>
				</xsl:if>
				<xsl:apply-templates select="atom:content/meta:properties[starts-with (data:ContentTypeId, $Waste-Items-Feed-In.Product-ContentTypeID)]">
					<xsl:with-param name="item" select="$item" />
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Property [@Name='Material']">
		<xsl:apply-templates select="$Waste-Items-Feed-In.Products [atom:content/meta:properties/data:Material = current()]" />
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="Object [@Type='System.Management.Automation.PSCustomObject']">
		<xsl:apply-templates select="Property" />
	</xsl:template>
	<!--
		Entry point matching to pwsh XML converted from CSV format
		Line 1: Material
		Line 2: <item #1 Material code>
		Line 3: <item #2 Material code>
		...
		[Xml] $local:items = $(Import-Csv -Delimiter ";" -Path $local:source | ConvertTo-Xml)
	-->
	<xsl:template match="/Objects">
		<List>
			<xsl:apply-templates select="Object" />
		</List>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Key" />
	<!--

	-->
	<xsl:template match="Material | Kurztext | Item">
		<xsl:copy>
			<xsl:apply-templates />
		</xsl:copy>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Entry">
		<xsl:copy>
			<xsl:apply-templates select="child::node()" />
			<Links>
				<xsl:value-of select="count (key ('items', Key))" />
			</Links>
		</xsl:copy>
	</xsl:template>
	<!--

	-->
	<xsl:template match="/List">
		<Root>
			<xsl:apply-templates select="Entry [generate-id (.) = generate-id (key ('items', Key)[1])]" >
				<xsl:sort select="Material" data-type="text" case-order="lower-first" order="ascending" />
			</xsl:apply-templates>
		</Root>
	</xsl:template>
	<!--

	-->
</xsl:stylesheet>
