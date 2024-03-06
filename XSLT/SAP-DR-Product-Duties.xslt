<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2021-2022)
		
	Revision:
		2021-02-21:	Erstellt.
		2021-04-19:	Verarbeitung von Bundle Items.
		2021-04-28:	Mode 'duty' bei der Verarbeitung der Elemente 'data:Duty_x002d_ListId' entfernt.
		2021-05-03:	Support for field data:Rank in Duties added.
		2021-05-22:	Notified-Products included in processed items.
		2022-06-05:	Changed to use hierarchical layout of ContentTypeID.
		2022-10-02:	Major change to support products in lots.
		2023-07-04:	Wildcard for BATCH selection added.
	Purpose:
		Flatens the items hierarchy such that all products are on the same level 
		and have all duties from the entire tree of child-items together with the 
		data of the linked device/battery/electric/chemical directly attached to.
-->
<xsl:stylesheet 
	version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:atom="http://www.w3.org/2005/Atom"
	xmlns:data="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:meta="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties" 
	exclude-result-prefixes="atom data meta georss gml">
	<!--

	-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--
		
	-->
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Duties.Batch" select="'*'"/>
	<xsl:param name="SAP-DR-Product-Duties.DutyLoadFile" />
	<xsl:param name="SAP-DR-Product-Duties.Product-ContentTypeID" select="'0x01003FAF714C6769BF4FA1B36DCF47ED659702'" />
	<!--
		
	-->
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry"/>
	<xsl:key name="dref" use="." match="/atom:feed/atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element"/>
	<xsl:key name="duty" use="data:Id" match="/atom:feed/atom:entry/atom:content/meta:properties"/>
	<!--
		Das Template kopiert, soweit es nicht durch das nachstehende verdraengt ist,
		die Daten der Produkte ohne den Namensraum. Es wird aufgerufen, wenn die
		Rekursion an einer 'duty' endet, deshalb ist der Modus 'final'.
	-->	
	<xsl:template match="data:*" mode="final">
		<xsl:element name="{local-name()}">
			<xsl:apply-templates />
		</xsl:element>
	</xsl:template>
	<!--
		Template leer fuer alle Elemente, die nicht in die Produkte kopiert werden sollen, 
		also alle systemdefinierten Elemente und die Referenzen auf items, parts, duties
	-->	
	<xsl:template match="data:REACH | data:Description1 | data:Material | data:ItemListId | data:Duty_x002d_ListId | data:Part_x002d_ListId | data:FileSystemObjectType | data:Id | data:ServerRedirectedEmbedUri | data:ServerRedirectedEmbedUrl | data:ID | data:ContentTypeId | data:Title | data:Modified | data:Created | data:AuthorId | data:EditorId | data:OData__UIVersionString | data:Attachments | data:GUID | data:ComplianceAssetId" mode="final" />
	<!--
		
	-->	
	<xsl:template match="data:element" mode="duty">
		<xsl:param name="entry"/>
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:variable name="tmp" select="." />
		<!--
			Der erste Test ist notwendig, weil auch Produkte in die Zwischendatei
			geschrieben werden, die keine 'duties' im aktuellen Batch haben.
		-->
		<xsl:if test="$duties and . = $duties">
			<xsl:if test="$verbose">
				<xsl:comment>
					<xsl:text>
				Recycling Duty:
				===========
					
					Origin (material where it is directly attached to):
						- Material:	</xsl:text><xsl:value-of select="$entry/atom:content/meta:properties/data:Material" /><xsl:text>
						- Description:	</xsl:text><xsl:value-of select="$entry/atom:content/meta:properties/data:Description1" /><xsl:if test="$entry/atom:content/meta:properties/data:Pieces [not (@meta:null)]"><xsl:text>
						- Lot size:		</xsl:text><xsl:value-of select="$entry/atom:content/meta:properties/data:Pieces" /><xsl:text> (order quantity will be diveded by lot size)</xsl:text></xsl:if><xsl:text>
						- Created:		</xsl:text><xsl:value-of select="$entry/atom:content/meta:properties/data:Modified" /><xsl:text>
						
					Details:
						- </xsl:text><xsl:value-of select="$duties[. = current()]/parent::meta:properties/data:Duty" /><xsl:text>
						- </xsl:text><xsl:value-of select="$duties[. = current()]/parent::meta:properties/data:Description12" /><xsl:text>
						- </xsl:text><xsl:value-of select="$duties[. = current()]/parent::meta:properties/data:Batch" /><xsl:text>
						
					Chain of recursively inspected items (Sharepoint IDs):
						- </xsl:text><xsl:value-of select="$loop" /><xsl:text>
						
		</xsl:text>
				</xsl:comment>
			</xsl:if>
			<!--
				
			-->
			<prd:duty SPKey="{.}">
				<prd:data>
					<xsl:variable name="base" select="$entry/atom:content/meta:properties" />
					<xsl:apply-templates select="$base/child::node() [not(@meta:null = 'true')]" mode="final" />
				</prd:data>
			</prd:duty>
		</xsl:if>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:element" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="key('entry',.)" mode="dig">
			<xsl:with-param name="duties" select="$duties" />
			<xsl:with-param name="loop" select="$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:element">
		<com:element>
			<xsl:value-of select="."/>
		</com:element>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:Duty_x002d_ListId">
		<xsl:param name="entry"/>
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="data:element" mode="duty">
			<xsl:with-param name="entry" select="$entry"/>
			<xsl:with-param name="duties" select="$duties" />
			<xsl:with-param name="loop" select="$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:apply-templates select="data:element" mode="dig">
			<xsl:with-param name="duties" select="$duties" />
			<xsl:with-param name="loop" select="$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:ItemListId">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<prd:duties>
			<xsl:apply-templates select="data:element" mode="dig">
				<xsl:with-param name="duties" select="$duties" />
				<xsl:with-param name="loop" select="$loop" />
			</xsl:apply-templates>
		</prd:duties>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:param name="count" />
		<xsl:if test="$count > 1">
			<xsl:apply-templates select=".">
				<xsl:with-param name="duties" select="$duties" />
				<xsl:with-param name="loop" select="$loop" />
				<xsl:with-param name="count" select="$count - 1" />
			</xsl:apply-templates>
		</xsl:if>
		<xsl:apply-templates select="key('entry',.)" mode="dig">
			<xsl:with-param name="duties" select="$duties" />
			<xsl:with-param name="loop" select="$loop" />
		</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template match="data:Country">
		<com:countries>
			<xsl:apply-templates select="data:element"/>
		</com:countries>
	</xsl:template>
	<!--
			Das nachstehende Template wird nur in der rekursiven Aufloesung
			der Produktstruktur angesprochen, wenn ein Verweis aus einer
			Stueck- oder Teileliste nachverfolgt wird. Es ist wichtig zu sehen,
			dass das Template auch auf die Knoten der 'duty' Definitionen
			passt. Dass es nicht mit solchen Knoten angesprochen wird, liegt
			an den Aufrufern.
	-->
	<xsl:template match="atom:entry" mode="dig">
		<xsl:param name="duties" />
		<xsl:param name="loop" />
		<xsl:choose>
			<xsl:when test="contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message>
					<xsl:text>[FATAL] SAP-DR-Product-Duties.xslt (line 163): Loop detected in chain </xsl:text><xsl:value-of select="$loop" /><xsl:text> Skipping.</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId [not(@meta:null = 'true')]">
					<xsl:with-param name="duties" select="$duties" />
					<xsl:with-param name="loop" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
					<xsl:with-param name="count" select="atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:ItemListId" mode="dig">
					<xsl:with-param name="duties" select="$duties" />
					<xsl:with-param name="loop" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId" mode="dig">
					<xsl:with-param name="duties" select="$duties" />
					<xsl:with-param name="loop" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Duty_x002d_ListId">
					<xsl:with-param name="duties" select="$duties" />
					<xsl:with-param name="loop" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
					<xsl:with-param name="entry" select="."/>
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--

	-->
	<xsl:template match="atom:content/meta:properties/data:Material">
		<prd:material>
			<xsl:value-of select="." />
		</prd:material>
	</xsl:template>
	<!--

	-->
	<xsl:template match="atom:content/meta:properties/data:Pieces [not (@meta:null)]">
		<prd:lotsize>
			<xsl:value-of select="." />
		</prd:lotsize>
	</xsl:template>
	<!--
			Das nachstehende Template arbeitet nur auf den Knoten der Produkt Definitionen,
			die sog. 'Products' sind, also Katalog-Eintraege. Die anderen Knoten der Produkt
			Definitionen werden nie an dieses Template geliefert.
	-->
	<xsl:template match="atom:entry">
		<xsl:param name="duties" />
		<xsl:if test="$verbose">
			<xsl:comment>
				<xsl:text>
		##########################################################################################
		##########################################################################################
										
					Catalogue product:
					==============
						- Material:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Material" /><xsl:text>
						- Description:	</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Description1" /><xsl:if test="atom:content/meta:properties/data:Pieces [not (@meta:null)]"><xsl:text>
						- Lot size:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Pieces" /><xsl:text> (order quantity will be diveded by lot size)</xsl:text></xsl:if><xsl:text>
						- Created:		</xsl:text><xsl:value-of select="atom:content/meta:properties/data:Modified" /><xsl:text>
						
	</xsl:text>
			</xsl:comment>
		</xsl:if>
		<prd:product SPKey="{atom:content/meta:properties/data:Id}">
			<xsl:apply-templates select="atom:content/meta:properties/data:Material" />
			<xsl:apply-templates select="atom:content/meta:properties/data:Pieces" />
			<!--
					Jetzt rekursiv die Produktstruktur abarbeiten ...
					duties:	Die Menge der 
			-->
			<xsl:apply-templates select="atom:content/meta:properties/data:ItemListId">
				<xsl:with-param name="duties" select="$duties" />
				<xsl:with-param name="loop" select="concat('[', atom:content/meta:properties/data:Id, ']')" />
			</xsl:apply-templates>
		</prd:product>
	</xsl:template>
	<!--
			Das Template erzeugt das Duty Verzeichnis am Anfang der Datei.
			Damit es keine Knoten der Produkt Definitionen verarbeitet, ist der
			Modus gegen die der Rekursion (., dig, duty, final) abgegrenzt.
	-->
	<xsl:template match="atom:entry" mode="init">
		<dty:entry SPKey="{atom:content/meta:properties/data:Id}">
			<dty:code>
				<xsl:value-of select="atom:content/meta:properties/data:Duty"/>
			</dty:code>
			<dty:batch>
				<xsl:value-of select="atom:content/meta:properties/data:Batch"/>
			</dty:batch>
			<dty:label>
				<xsl:value-of select="atom:content/meta:properties/data:Description12"/>
			</dty:label>
			<dty:rank>
				<xsl:value-of select="atom:content/meta:properties/data:Rank"/>
			</dty:rank>
			<xsl:apply-templates select="atom:content/meta:properties/data:Country"/>
		</dty:entry>
	</xsl:template>
	<!--
			Eingangstemplate, passt auf den Wurzelknoten der Produkt-Definitionen.
			
			drefs:	Die Menge aller SPKeys in Referenzen der Produkte auf 'duties'
			used: 	Alle 'duty' Definitionen, auf die in der geladenen Produktmenge
						verwiesen wird, d.h. deren SPKey in drefs enthalten ist.
	-->
	<xsl:template match="/atom:feed">
		<xsl:variable name="drefs" select="atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element"/>
		<xsl:variable name="used" select="document($SAP-DR-Product-Duties.DutyLoadFile)/atom:feed/atom:entry[atom:content/meta:properties/data:Id = $drefs]['*' = $SAP-DR-Product-Duties.Batch or atom:content/meta:properties/data:Batch = $SAP-DR-Product-Duties.Batch]"/>
		<xsl:comment>
			<xsl:text>
	++++++++++++++++++++++++++++++++++++++++++++++++++++
	
		Debug: </xsl:text><xsl:value-of select="$debug" />
			<xsl:text>
			
		Verbose: </xsl:text><xsl:value-of select="$verbose" />
			<xsl:text>
			
	++++++++++++++++++++++++++++++++++++++++++++++++++++		
</xsl:text>
		</xsl:comment>
		<prd:root batch="{$SAP-DR-Product-Duties.Batch}">
			<dty:references>
				<xsl:attribute name="SPKeys">
					<xsl:for-each select="$used/atom:content/meta:properties/data:Id">
						<xsl:value-of select="."/>
							<xsl:if test="position () != last ()">
								<xsl:text>,</xsl:text>
							</xsl:if>
						</xsl:for-each>
					</xsl:attribute>
				<!--
						Nur Duties, die den Stammdaten durch Items refernziert sind, werden
						in die temoraere Datei mit der flachen Hierarchie geschrieben. Der Modus
						ist explizit 'init', damit die Verarbeitung der Duty-Definitionen gegen die
						der Produkt-Definitionen abgegrenzt ist.
				-->
				<xsl:apply-templates select="$used" mode="init"/>
				<!--

				-->
			</dty:references>
			<!--
				Der nachfolgende Templateanwendung betrifft nur Produkte: Dazu muss
				allerdings in der aufrufenden Datei der Parameter auf die ContentTypeID
				der Produkte auf dem SharePoint passen. Hier wird der hierarchische
				Aufbau der ContentTypeID genutzt, der auf dem iptrack SharePoint ist:
				Item						0x01003FAF714C6769BF4FA1B36DCF47ED6597
				Product					0x01003FAF714C6769BF4FA1B36DCF47ED659702
				Notified_Product	0x01003FAF714C6769BF4FA1B36DCF47ED65970201
				Product-in-Lots		0x01003FAF714C6769BF4FA1B36DCF47ED65970202
			-->
			<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, $SAP-DR-Product-Duties.Product-ContentTypeID)]">
				<xsl:with-param name="duties" select="$used/atom:content/meta:properties/data:Id" />
			</xsl:apply-templates>
			<!--
				
			-->
		</prd:root>
	</xsl:template>
	<!--

	-->
</xsl:stylesheet>
