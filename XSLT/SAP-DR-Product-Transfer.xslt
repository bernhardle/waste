<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2021-2023)
		
	Revision:
		2021-11-14:	Created.
		2022-10-02:	Major revision tag to reflect readiness for products in lots.
		2023-01-05:	Transfer targets parameterized.
-->
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	exclude-result-prefixes="atom com data meta">
	<!--

	-->
	<xsl:output method="xml" encoding="utf-8" indent="yes" standalone="yes" />
	<!--

	-->
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Transfer.ContentType" />
	<xsl:param name="SAP-DR-Product-Transfer.ProductLoadFile" />
	<xsl:param name="SAP-DR-Product-Transfer.ContentTypesLoadFile" />
	<!--
		Herunterladen der Dateien mit:
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/ContentTypes
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/items?$select=Id,ContentTypeId,Material,Description1,ItemListId,Reference_x002d_ProductId,Part_x002d_ListId,Pieces&$top=100000&$order=ID&$filter=ContentTypeId eq '0x01003FAF714C6769BF4FA1B36DCF47ED659702009768DCF9E8E5424EB2E68CAB20681245' or ContentTypeId eq '0x01003FAF714C6769BF4FA1B36DCF47ED659702010044799B555B8A22498614585A51717B86' or ContentTypeId eq '0x01003FAF714C6769BF4FA1B36DCF47ED65970400F6E244CC367A26439227A016D1DF32CD'
	-->
	<xsl:variable name="SAP-DR-Product-Transfer.Products" select="document($SAP-DR-Product-Transfer.ProductLoadFile)/atom:feed/atom:entry" />
	<xsl:variable name="SAP-DR-Product-Transfer.ContentTypes" select="document($SAP-DR-Product-Transfer.ContentTypesLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<!--

	-->
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry" />
	<xsl:key name="item" use="com:material" match="com:item" />
	<xsl:key name="base" use="com:base-key" match="com:item" />
	<xsl:key name="base-my" use="concat (com:base-key,'-',com:my-key)" match="com:item" />
	<!--
		
	-->
	<xsl:template match="data:element">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="pieces" select="$pieces" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="data:element">
			<xsl:with-param name="pieces" select="$pieces" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="pieces" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="pieces" select="$pieces" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--
			Das nachstehende Template wird nur in der rekursiven Auflösung
			der Produktstruktur angesprochen, wenn ein Verweis aus einer
			Stueck- oder Teileliste nachverfolgt wird.
	-->
	<xsl:template match="atom:entry">
		<xsl:param name="pieces" />
		<xsl:param name="loop" select="'>>'" />
		<xsl:param name="base" select="." />
		<xsl:variable name="next" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
		<xsl:variable name="cotyid" select="atom:content/meta:properties/data:ContentTypeId" />
		<xsl:variable name="content" select="$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name" />
		<xsl:choose>
			<xsl:when test="contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Transfer.xslt (line 98): Loop detected in chain </xsl:text><xsl:value-of select="$next" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<!--
					
				-->
				<xsl:if test="$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name = $SAP-DR-Product-Transfer.ContentType">
					<com:item>
						<com:base-key>
							<xsl:value-of select="$base/atom:content/meta:properties/data:ID" />
						</com:base-key>
						<com:my-key>
							<xsl:value-of select="atom:content/meta:properties/data:ID" />
						</com:my-key>
						<com:material>
							<xsl:value-of select="atom:content/meta:properties/data:Material" />
						</com:material>
						<com:description>
							<xsl:value-of select="atom:content/meta:properties/data:Description1" />
						</com:description>
						<com:amount>
							<xsl:value-of select="$pieces" />
						</com:amount>
					</com:item>
				</xsl:if>
				<!--
					
				-->
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId">
					<xsl:with-param name="loop" select="$next" />
					<xsl:with-param name="base" select="$base" />
					<xsl:with-param name="pieces" select="$pieces * atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<!--
					
				-->
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId | atom:content/meta:properties/data:ItemListId">
					<xsl:with-param name="loop" select="$next" />
					<xsl:with-param name="base" select="$base" />
					<xsl:with-param name="pieces">
						<xsl:choose>
							<xsl:when test="$SAP-DR-Product-Transfer.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name = 'Product-in-Lots'">
								<xsl:value-of select="$pieces div atom:content/meta:properties/data:Pieces" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$pieces" />
							</xsl:otherwise>
						</xsl:choose>					
					</xsl:with-param>
				</xsl:apply-templates>
				<!--
					
				-->
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="Object [Property/@Name='Material' and Property/@Name='Anzahl']">
			<xsl:apply-templates select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:Material = current()/Property[@Name='Material']]">
				<xsl:with-param name="pieces" select="current()/Property[@Name='Anzahl']" />
			</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Objects">
		<com:items>
			<xsl:apply-templates select="Object" />
		</com:items>
	</xsl:template>
	<!--
		$(Get-Content ..\Downloads\Transfer.xml).Root.ChildNodes | Select-Object -Property Material,@{Name="Amount"; Expression={$_.Anzahl -as [Int]}} | Out-GridView -Title "Products transferred to Packaging"
	-->
	<xsl:template match="/com:items" mode="none">
		<Root>
			<xsl:for-each select="com:item [generate-id (.) = generate-id (key ('item', com:material)[1])]">
				<xsl:sort select="com:material" />
				<Item>
					<Material>
						<xsl:value-of select="com:material" />
					</Material>
					<Materialkurztext>
						<xsl:value-of select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:Material = current()/com:material]/atom:content/meta:properties/data:Description1" />
					</Materialkurztext>
					<Anzahl>
						<xsl:value-of select="sum (key ('item', com:material)/com:amount)" />
					</Anzahl>
				</Item>
			</xsl:for-each>
		</Root>
	</xsl:template>
	<!--

	-->
	<xsl:template name="rec-sum-sel">
		<xsl:param name="items" />
		<xsl:param name="count" select="0" />
		<xsl:choose>
			<xsl:when test="$items [position () > 1]">
				<xsl:call-template name="rec-sum-sel">
					<xsl:with-param name="items" select="$items [position () > 1]" />
					<xsl:with-param name="count" select="$count + com:amount" />
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<Base>
					<xsl:value-of select="$items [1]/com:base-key" />
				</Base>
				<Material>
					<xsl:value-of select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:ID = $items [1]/com:base-key]/atom:content/meta:properties/data:Material" />
				</Material>
				<Product>
					<xsl:value-of select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:ID = $items [1]/com:base-key]/atom:content/meta:properties/data:Description1" />
				</Product>
				<Key>
					<xsl:value-of select="$items [1]/com:my-key" />
				</Key>
				<Battery>
					<xsl:value-of select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:ID = $items [1]/com:my-key]/atom:content/meta:properties/data:Description1" />
				</Battery>
				<Weight>
					<xsl:value-of select="$SAP-DR-Product-Transfer.Products [atom:content/meta:properties/data:ID = $items [1]/com:my-key]/atom:content/meta:properties/data:Weight" />
				</Weight>
				<Number>
					<xsl:value-of select="$items [1]/com:amount + $count" />
				</Number>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--

	-->
	<xsl:template match="/com:items">
		<Root>
			<xsl:for-each select="com:item [generate-id (.) = generate-id (key ('base-my', concat (com:base-key,'-',com:my-key))[1])]">
				<xsl:variable name="selection" select="key ('base-my', concat (current()/com:base-key,'-',current()/com:my-key))"/>
				<Item>
					<xsl:call-template name="rec-sum-sel">
						<xsl:with-param name="items" select="$selection" />
					</xsl:call-template>
				</Item>
			</xsl:for-each>
		</Root>
	</xsl:template>
	<!--
		
	-->
</xsl:stylesheet>
