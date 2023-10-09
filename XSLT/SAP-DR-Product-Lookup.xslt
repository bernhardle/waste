<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2021-2022)
		
	Revision:
		2021-04-30:	Created.
		2021-05-02:	Format of attachment value changed to checkbox like style.
		2021-05-22:	Added modes for back tracking and export of tags (WEEE/BATT/TVVV/SCIP).
		2021-05-31:	Added mode for duty lookup.
		2022-06-05:	Changed to use hierarchical layout of ContentTypeID.
		2022-10-02:	Major change ready to support products in lots.
		2023-01-04:	Encoding changed from "asci" to "UTF-8"
		2023-10-09: [Linux] Removed case-order from xsl:sort
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
	<xsl:output method="text" encoding="UTF-8" />
	<!--

	-->
	<xsl:param name="verbose" select="false()" />
	<xsl:param name="debug" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.Mode" select="'forward'" />
	<xsl:param name="SAP-DR-Product-Lookup.UnfoldBundles" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.ShowDuties" select="false()" />
	<xsl:param name="SAP-DR-Product-Lookup.ProductLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.DutyLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.FieldsLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.ContentTypesLoadFile" />
	<xsl:param name="SAP-DR-Product-Lookup.Material" />
	<xsl:param name="SAP-DR-Product-Lookup.Duty" />
	<xsl:param name="SAP-DR-Product-Lookup.Product-ContentTypeID" select="'0x01003FAF714C6769BF4FA1B36DCF47ED659702'" />
	<!--
		Herunterladen der Dateien mit:
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/ContentTypes
			https://iptrack.sharepoint.com/sites/RESTAPI/_api/Web/Lists(guid'5cd75f66-486f-49fa-8176-b3e74fc8a10d')/Fields?$select=InternalName,Title,Description,TypeAsString,TypeDisplayName,TypeShortDescription
	-->
	<xsl:variable name="SAP-DR-Product-Lookup.Products" select="document($SAP-DR-Product-Lookup.ProductLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.Duties" select="document($SAP-DR-Product-Lookup.DutyLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.Fields" select="document($SAP-DR-Product-Lookup.FieldsLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<xsl:variable name="SAP-DR-Product-Lookup.ContentTypes" select="document($SAP-DR-Product-Lookup.ContentTypesLoadFile)/atom:feed/atom:entry/atom:content/meta:properties" />
	<!--

	-->
	<xsl:variable name="SAP-DR-Product-Lookup.ConstBlank140" select="'                                                                                                                                            '"/>
	<xsl:variable name="SAP-DR-Product-Lookup.ConstHyphen140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'"/>
	<!--

	-->
	<xsl:key name="entry" use="atom:content/meta:properties/data:Id" match="/atom:feed/atom:entry" />
	<xsl:key name="back" use="atom:content/meta:properties/data:ItemListId/data:element | atom:content/meta:properties/data:Part_x002d_ListId/data:element | atom:content/meta:properties/data:Reference_x002d_ProductId" match="/atom:feed/atom:entry" />
	<xsl:key name="dref" use="." match="/atom:feed/atom:entry/atom:content/meta:properties/data:Duty_x002d_ListId/data:element" />
	<xsl:key name="duty" use="data:Batch" match="/atom:feed/atom:entry/atom:content/meta:properties" />
	<!--
		Das Template kopiert, soweit es nicht durch das nachstehende verdraengt ist,
		die Daten der Produkte ohne den Namensraum. Es wird aufgerufen, wenn die
		Rekursion an einer 'duty' endet, deshalb ist der Modus 'final'.
	-->	
	<xsl:template match="data:*" mode="final">
		<xsl:param name="level" />
		<xsl:variable name="name" select="local-name()" />
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="label" select="$SAP-DR-Product-Lookup.Fields [data:InternalName = $name]/data:Title" />
			<xsl:with-param name="value" select="." />
		</xsl:call-template>
	</xsl:template>
	<!--
		Template leer fuer alle Elemente, die nicht in die Produkte kopiert werden sollen, 
		also alle systemdefinierten Elemente und die Referenzen auf items, parts, duties
	-->	
	<xsl:template match="data:Description1 | data:Material | data:ItemListId | data:Duty_x002d_ListId | data:Part_x002d_ListId | data:Reference_x002d_ProductId | data:FileSystemObjectType | data:Id | data:ServerRedirectedEmbedUri | data:ServerRedirectedEmbedUrl | data:ID | data:ContentTypeId | data:Title | data:Modified | data:Created | data:AuthorId | data:EditorId | data:OData__UIVersionString | data:Attachments | data:GUID | data:ComplianceAssetId" mode="final" />
	<!--
		Modus 'probe' zum Austesten der Duties auf den nachgelagerten Items
	-->
	<xsl:template match="data:element">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--
		Template gibt die Duties geordnet nach dem Batch (BATT, TVVV, WEEE, VOID) aus.
	-->
	<xsl:template match="data:Duty_x002d_ListId">
		<xsl:param name="level" />
		<xsl:param name="entry"/>
		<xsl:variable name= "tmp" select="$SAP-DR-Product-Lookup.Duties [data:Id = current()/data:element]" />
		<xsl:if test="$tmp">
			<xsl:for-each select="$tmp [generate-id(.) = generate-id (key ('duty', data:Batch)[data:Id = current()/data:element][1])]">
				<xsl:call-template name="line-fill">
					<xsl:with-param name="value" select="data:Batch" />
					<xsl:with-param name="level" select="$level" />
				</xsl:call-template>
				<xsl:for-each select="$tmp [data:Batch = current()/data:Batch]">
					<xsl:sort select="data:Duty" data-type="text"  order="ascending" />
					<xsl:call-template name="line-fill">
						<xsl:with-param name="level" select="$level" />
						<xsl:with-param name="label" select="data:Duty" />
						<xsl:with-param name="value" select="data:Description12" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:for-each>
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="$level" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="data:Part_x002d_ListId | data:ItemListId">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:apply-templates select="data:element">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="data:Reference_x002d_ProductId">
		<xsl:param name="level" />
		<xsl:param name="loop" />
		<xsl:param name="base" />
		<xsl:param name="count" />
		<xsl:if test="$SAP-DR-Product-Lookup.UnfoldBundles and $count > 1">
			<xsl:apply-templates select=".">
				<xsl:with-param name="level" select="$level" />
				<xsl:with-param name="loop" select="$loop" />
				<xsl:with-param name="base" select="$base" />
				<xsl:with-param name="count" select="$count - 1" />
			</xsl:apply-templates>
		</xsl:if>
		<xsl:apply-templates select="key('entry',.)">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="loop" select="$loop" />
			<xsl:with-param name="base" select="$base" />
		</xsl:apply-templates>
	</xsl:template>
	<!--

	-->
	<xsl:template name="line-fill">
		<xsl:param name="label" select="''"/>
		<xsl:param name="value" select="''"/>
		<xsl:param name="level" />
		<xsl:variable name="border" select="'##'" />
		<xsl:variable name="margin" select="2" />
		<xsl:variable name="indent" select="8" />
		<xsl:variable name="header" select="12" />
		<xsl:variable name="width" select="100" />
		<xsl:variable name="nlabel" select="substring ($label, 1, $header)" />
		<xsl:choose>
			<xsl:when test="string-length ($label) = 0 and string-length ($value) = 0">
				<xsl:value-of select="concat ($border, substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $margin), substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $indent * $level), '+', substring ($SAP-DR-Product-Lookup.ConstHyphen140, 1, $width - ($indent * $level)),'+&#xA;')" />
			</xsl:when>
			<xsl:when test="string-length ($label) = 0">
				<xsl:variable name="nvalue" select="substring ($value, 1, $width - $indent * $level - 2)" />
				<xsl:value-of select="concat ($border, substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $margin), substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $indent * $level), '| ', $nvalue, substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $width - ($indent * $level) - string-length ($nvalue) - 1), '|&#xA;')" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="nvalue" select="substring ($value, 1, $width - $indent * $level - $header - 7)" />
				<xsl:value-of select="concat ($border, substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $margin), substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $indent * $level), '|  - ', $nlabel, ': ', substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $header - string-length ($nlabel)), $nvalue, substring ($SAP-DR-Product-Lookup.ConstBlank140, 1, $width - ($indent * $level) - $header - string-length ($nvalue) - 6), '|&#xA;')" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		Template untersucht rekursiv alle nachgelagerten Items auf Duties im 
		angegeben Batch und gibt ggf. eine 1 zurück. Die Verarbeitung wird
		am ersten Treffer abgebrochen. Durchlauf bis zum Ende ohne Treffer
		liefert eine 0.
	-->
	<xsl:template name="rec-probe-duties">
		<xsl:param name="loop" />
		<xsl:param name="batch" />
		<xsl:param name="duty" />
		<xsl:param name="items" />
		<xsl:choose>
			<xsl:when test="count ($SAP-DR-Product-Lookup.Duties[data:Batch = $batch or data:Id = $duty][data:Id = $items [1]/atom:content/meta:properties/data:Duty_x002d_ListId/data:element]) > 0">
				<xsl:value-of select="1" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="keys" select="$items [1]/atom:content/meta:properties/data:Reference_x002d_ProductId | $items [1]/atom:content/meta:properties/data:ItemListId/data:element | $items[1]/atom:content/meta:properties/data:Part_x002d_ListId/data:element" />
				<xsl:if test="$keys">
				<xsl:variable name="tmp">
					<xsl:call-template name="rec-probe-duties">
						<xsl:with-param name="items" select="key('entry', $keys)" />
						<xsl:with-param name="loop" select="concat($loop, '-[', $items [1]/atom:content/meta:properties/data:Id, ']')" />
						<xsl:with-param name="batch" select="$batch" />
						<xsl:with-param name="duty" select="$duty" />
					</xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$tmp > 0">
						<xsl:value-of select="1" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$items [position () > 1]">
								<xsl:call-template name="rec-probe-duties">
									<xsl:with-param name="loop"  select="$loop"/>
									<xsl:with-param name="batch" select="$batch"/>
									<xsl:with-param name="duty" select="$duty" />
									<xsl:with-param name="items" select="$items [position () > 1]" />
								</xsl:call-template>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="0" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="atom:entry" mode="bernie">
		<xsl:param name="level" />
		<xsl:param name="back" />
		<xsl:param name="base" />
		<xsl:variable name="cotyid" select="atom:content/meta:properties/data:ContentTypeId" />
		<xsl:if test="$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name = 'Sales-Packaging'">
			<xsl:value-of select="concat($base/atom:content/meta:properties/data:Material,' - ', atom:content/meta:properties/data:Material, '&#xA;')" />
		</xsl:if>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="atom:entry" mode="text">
		<xsl:param name="level" />
		<xsl:param name="back" select="false ()" />
		<xsl:variable name="shift" select="10" />
		<xsl:variable name="cotyid" select="atom:content/meta:properties/data:ContentTypeId" />
		<!--
			Rahmenlinie oben
		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
		</xsl:call-template>
		<!--
			Kopfzeile
		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="value">
				<xsl:value-of select="$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name" />
				<xsl:if test="string-length($SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Description) > 0">
					<xsl:value-of select="concat (' (', $SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Description,')')" />
				</xsl:if>
			</xsl:with-param>
		</xsl:call-template>
		<!--
			Materialkurztext
		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="label" select="'Kurztext'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Description1" />
		</xsl:call-template>
		<!--
			Material
		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="label" select="'Material'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Material" />
		</xsl:call-template>
		<!--

		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="label" select="'Attachments'" />
			<xsl:with-param name="value">
				<xsl:choose>
					<xsl:when test="atom:content/meta:properties/data:Attachments = 'true'">
						<xsl:text>[X]</xsl:text>
					</xsl:when>
					<xsl:otherwise>
						<xsl:text>[_]</xsl:text>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:with-param>
		</xsl:call-template>
		<!--
			Erstelldatum
		-->
		<xsl:call-template name="line-fill">
			<xsl:with-param name="level" select="$level" />
			<xsl:with-param name="label" select="'Erstellt'" />
			<xsl:with-param name="value" select="atom:content/meta:properties/data:Modified" />
		</xsl:call-template>
		<!--
			Sonstige nichtausgeblendete Datenfelder
		-->
		<xsl:apply-templates select="atom:content/meta:properties/data:* [not(@meta:null = 'true')]" mode="final">
			<xsl:with-param name="level" select="$level" />
		</xsl:apply-templates>
		<!--
			Recycling Tags in SAP
		-->
		<xsl:if test="$SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name = 'Product' or $SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $cotyid]/data:Name = 'Notified-Product'">
			<xsl:variable name="batt">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'BATT'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="tvvv">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'TVVV'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="weee">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'WEEE'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:variable name="scip">
				<xsl:call-template name="rec-probe-duties">
					<xsl:with-param name="items" select="." />
					<xsl:with-param name="batch" select="'SCIP'" />
					<xsl:with-param name="level" select="0" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="$level" />
				<xsl:with-param name="value" select="concat ('   BATT [', translate($batt, '01', '_X'), ']      TVVV [', translate($tvvv, '01', '_X'), ']      WEEE [', translate($weee, '01', '_X'), ']      SCIP [', translate($scip, '01', '_X'),']')" />
			</xsl:call-template>
		</xsl:if>
		<!--
			Rahmenlinie unten, falls keine Duties angezeigt werden
		-->
		<xsl:if test="$back or not($SAP-DR-Product-Lookup.ShowDuties and atom:content/meta:properties/data:Duty_x002d_ListId/data:element)">
			<xsl:call-template name="line-fill">
				<xsl:with-param name="level" select="$level" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<!--
			Das nachstehende Template wird nur in der rekursiven Aufloesung
			der Produktstruktur angesprochen, wenn ein Verweis aus einer
			Stueck- oder Teileliste nachverfolgt wird. Es ist wichtig zu sehen,
			dass das Template auch auf die Knoten der 'duty' Definitionen
			passt. Dass es nicht mit solchen Knoten angesprochen wird, liegt
			an den Aufrufern.
	-->
	<xsl:template match="atom:entry">
		<xsl:param name="level" />
		<xsl:param name="indent" />
		<xsl:param name="loop" select="'>>'" />
		<xsl:param name="base" select="." />
		<xsl:param name="trace" select="''" />
		<xsl:variable name="next" select="concat($loop, '-[', atom:content/meta:properties/data:Id, ']')" />
		<xsl:choose>
			<xsl:when test="contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 386): Loop detected in chain </xsl:text><xsl:value-of select="$next" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<!--
					
				-->
				<xsl:choose>
					<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'attachments'">
						<xsl:variable name="props" select="atom:content/meta:properties" />
						<xsl:if test="not (contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))) and atom:content/meta:properties/data:Attachments = 'true'">
							<xsl:value-of select="concat ($base/atom:content/meta:properties/data:Id, ';', translate ($base/atom:content/meta:properties/data:Material, ';',','), ';', translate ($base/atom:content/meta:properties/data:Description1, ';', ','), ';', $props/data:Id, ';', translate ($props/data:Material, ';[](){}|', ','), ';', $SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = $props/data:ContentTypeId]/data:Name, ';', translate ($props/data:Description1, ';', ','), '&#xA;')" />
						</xsl:if>
					</xsl:when>
					<!--
						
					-->
					<xsl:otherwise>
						<xsl:apply-templates select="." mode="text">
							<xsl:with-param name="level" select="$level" />
						</xsl:apply-templates>
						<xsl:if test="$SAP-DR-Product-Lookup.ShowDuties">
							<xsl:apply-templates select="atom:content/meta:properties/data:Duty_x002d_ListId">
								<xsl:with-param name="level" select="$level" />
								<xsl:with-param name="entry" select="."/>
							</xsl:apply-templates>
						</xsl:if>
					</xsl:otherwise>
				</xsl:choose>
				<!--
					
				-->
				<xsl:apply-templates select="atom:content/meta:properties/data:Reference_x002d_ProductId">
					<xsl:with-param name="level" select="$level + 1" />
					<xsl:with-param name="loop" select="$next" />
					<xsl:with-param name="base" select="$base" />
					<xsl:with-param name="count" select="atom:content/meta:properties/data:Pieces" />
				</xsl:apply-templates>
				<xsl:apply-templates select="atom:content/meta:properties/data:Part_x002d_ListId | atom:content/meta:properties/data:ItemListId">
					<xsl:with-param name="level" select="$level + 1" />
					<xsl:with-param name="loop" select="$next" />
					<xsl:with-param name="base" select="$base" />
				</xsl:apply-templates>
				<!--

				-->
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		Template wird aufgerufen im Modus 'back' auf einem Item ohne
		referenzierende andere Items. Typisch wird das ein Typ Product
		oder davon abgeleitetes sein. Das Template verfolgt dann die
		gespeicherte Suchkette bis zum Anfang und schreibt die Infos.
	-->
	<xsl:template name="track-down">
		<xsl:param name="level" select="0" />
		<xsl:param name="chain" />
		<xsl:variable name="key" select="substring-before (substring-after ($chain, '['), ']')" />
		<xsl:if test="not (contains($key, 'x'))">
			<xsl:apply-templates select="key ('entry', $key)" mode="text" >
				<xsl:with-param name="level" select="$level" />
				<xsl:with-param name="back" select="true ()" />
			</xsl:apply-templates>
			<xsl:call-template name="track-down">
				<xsl:with-param name="level" select="$level + 1" />
				<xsl:with-param name="chain" select="substring ($chain, string-length ($key) + 3)" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="atom:entry" mode="back">
		<xsl:param name="level" select="3" />
		<xsl:param name="loop" select="'[x]'" />
		<xsl:variable name="top" select="key ('back', atom:content/meta:properties/data:Id)" />
		<xsl:choose>
			<xsl:when test="contains($loop, concat ('[', atom:content/meta:properties/data:Id, ']'))">
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 461): Loop detected in chain </xsl:text><xsl:value-of select="concat('[', atom:content/meta:properties/data:Id, ']-', $loop)" /><xsl:text> Skipping.

</xsl:text>
				</xsl:message>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$top">
						<xsl:apply-templates select="$top" mode="back">
							<xsl:with-param name="level" select="$level - 1" />
							<xsl:with-param name="loop" select="concat('[', atom:content/meta:properties/data:Id, ']-', $loop)" />
						</xsl:apply-templates>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="track-down">
							<xsl:with-param name="chain" select="concat('[', atom:content/meta:properties/data:Id, ']-', $loop)" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--

	-->
	<xsl:template match="atom:entry" mode="tags">
		<xsl:variable name="batt">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'BATT'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="tvvv">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'TVVV'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="weee">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'WEEE'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="scip">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="batch" select="'SCIP'" />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select="concat (atom:content/meta:properties/data:Material, ';', translate(atom:content/meta:properties/data:Description1, ';',' '), ';' , $batt, ';', $tvvv, ';', $weee, ';', $scip, '&#xD;&#xA;')" />
	</xsl:template>
	<!--

	-->
	<xsl:template match="atom:entry" mode="duty">
		<xsl:variable name="hit">
			<xsl:call-template name="rec-probe-duties">
				<xsl:with-param name="duty" select="$SAP-DR-Product-Lookup.Duties [data:Duty = $SAP-DR-Product-Lookup.Duty]/data:Id" />
				<xsl:with-param name="items" select="." />
				<xsl:with-param name="level" select="0" />
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test="$hit = 1">
			<xsl:value-of select="concat (atom:content/meta:properties/data:Material, ';', translate(atom:content/meta:properties/data:Description1, ';',' '), ';&#xD;&#xA;')" />
		</xsl:if>
	</xsl:template>
	<!--
			Eingangstemplate, passt auf den Wurzelknoten der Produkt-Definitionen.
			Modi:
				-	attachments:	
				-	forward:			
				-	backward:		
				-	tagexport:
				-	duty:				
	-->
	<xsl:template match="/atom:feed">
		<xsl:choose>
			<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'attachments'">
				<xsl:value-of select="concat ('Base_Id', ';', 'Base_Material', ';', 'Base_Kurztext', ';', 'Item_Id', ';', 'Item_Material', ';', 'Item_Typ', ';', 'Item_Kurztext', '&#xA;')" />
				<xsl:apply-templates select="atom:entry [atom:content/meta:properties/data:Material = $SAP-DR-Product-Lookup.Material]">
					<xsl:with-param name="level" select="0" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'forward'">
				<xsl:apply-templates select="atom:entry [$SAP-DR-Product-Lookup.Material = '*' or atom:content/meta:properties/data:Material = $SAP-DR-Product-Lookup.Material]">
					<xsl:with-param name="level" select="0" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'backward'">
				<xsl:apply-templates select="atom:entry[atom:content/meta:properties/data:Material = $SAP-DR-Product-Lookup.Material]" mode="back" />
			</xsl:when>
			<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'tagexport'">
				<xsl:value-of select="'Material;Materialkurztext;BATT;TVVV;WEEE;SCIP&#xD;&#xA;'" />
				<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, $SAP-DR-Product-Lookup.Product-ContentTypeID)]" mode="tags">
					<xsl:sort data-type="text" lang="en" order="ascending" select="atom:content/meta:properties/data:Material" />
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="$SAP-DR-Product-Lookup.Mode = 'duty'">
				<xsl:variable name="tmp" select="$SAP-DR-Product-Lookup.Duties [data:Duty = $SAP-DR-Product-Lookup.Duty]" />
				<xsl:choose>
					<xsl:when test="count ($tmp) = 1">
						<xsl:value-of select="'Material;Materialkurztext;&#xD;&#xA;'" />
						<xsl:apply-templates select="atom:entry [starts-with (atom:content/meta:properties/data:ContentTypeId, $SAP-DR-Product-Lookup.Product-ContentTypeID)]" mode="duty">
							<xsl:sort data-type="text" lang="en" order="ascending" select="atom:content/meta:properties/data:Material" />
						</xsl:apply-templates>
					</xsl:when>
					<xsl:otherwise>
						<xsl:message terminate="no">
							<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 567): Invalid duty &apos;</xsl:text><xsl:value-of select="$SAP-DR-Product-Lookup.Duty" /><xsl:text>&apos;. Skipping.

</xsl:text>
						</xsl:message>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:message terminate="no">
					<xsl:text>
					
[FATAL] SAP-DR-Product-Lookup.xslt (line 578): Wrong mode &apos;</xsl:text><xsl:value-of select="$SAP-DR-Product-Lookup.Mode" /><xsl:text>&apos;. Skipping.

</xsl:text>
				</xsl:message>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Objects">
		<xsl:value-of select="'Material;Materialkurztext;Produktart&#xD;&#xA;'" />
		<xsl:for-each select="$SAP-DR-Product-Lookup.Products [starts-with (data:ContentTypeId, $SAP-DR-Product-Lookup.Product-ContentTypeID)]">
			<xsl:sort select="data:Material" data-type="text" order="ascending" />
			<xsl:value-of select="concat(data:Material, ';', data:Description1, ';', $SAP-DR-Product-Lookup.ContentTypes [data:Id/data:StringValue = current ()/data:ContentTypeId]/data:Name, '&#xD;&#xA;')" />
		</xsl:for-each>
	</xsl:template>
	<!--

	-->
</xsl:stylesheet>
