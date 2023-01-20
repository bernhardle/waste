<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2022-2023)
		
	Revision:
		2022-06-12:	Created.
		2022-10-02:	Major change to support products in lots.
		2023-01-06:	Three steps conversion for CSV based XML to deal with ill-formatted numbers.
		
	Function:
		Converts an XML 2003 encoded Excel spreadsheet list of
		sales volumes into an intermediate format. This is a mere
		structure transformation for most products with the
		exception of products which are required to be ordered
		in predetermined lot sizes > 1. These products are of the
		type Product-in-Lots and have all their respective individual 
		sales number (not the aggregated volume) diveded by the 
		predetermined lot size in transformation. Errors are flagged
		where an individual sales number is not an exact multiple of
		the predetermined lot size.
		
-->
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:ns1="http://www.w3.org/2005/Atom" 
	xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" 
	xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" 
	xmlns:georss="http://www.georss.org/georss" 
	xmlns:gml="http://www.opengis.net/gml" 
	xmlns:xls="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
	xmlns:v="urn:schemas-microsoft-com:vml" 
	xmlns:x="urn:schemas-microsoft-com:office:excel" 
	xmlns:xlink="http://www.w3.org/1999/xlink" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:msg="http://www.rothenberger.com/productcompliance/recycling/messages" 
	exclude-result-prefixes="com d georss gml m msg ns1 o prd ss v x xls xlink">
	<!--
		
	-->
	<xsl:param name="debug" select="0" />
	<xsl:param name="verbose" select="0" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.MasterData" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.Country" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.NumberGroupSeparator" select="'.'" />
	<xsl:param name="SAP-DR-Recycling-Preprocess.NumberDecimalSeparator" select="','" />
	<!--
		
	-->
	<xsl:variable name="master" select="document ($SAP-DR-Recycling-Preprocess.MasterData)/prd:root" />
	<xsl:variable name="product" select="$master/prd:product" />
	<!--
		Formatting settings should be the same as in SAP-DR-Recycling-Calculate.xslt
	-->
	<xsl:variable name="constBlk140" select="'                                                                                                                                            '" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="constDot140" select="'............................................................................................................................................'" />
	<xsl:variable name="blockWidth" select="95" />
	<xsl:variable name="indent1" select="10" />
	<xsl:variable name="indent2" select="22" />
	<xsl:variable name="indent3" select="45" />
	<xsl:variable name="indent4" select="70" />
	<!--
		
	-->
	<xsl:key name="mat-xml" use="concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data)" match="xls:Row" />
	<xsl:key name="mat-csv" use="concat (Property[@Name='Material'],'-', Property[@Name='Empfangsland'])" match="Object" />
	<!--
		
	-->
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<!--
		
	-->
	<xsl:output method="xml" encoding="UTF-8" indent="yes" standalone="yes"  version="1.0" />
	<!--
		**********************************************************************
			Das Template ...
			
		**********************************************************************
	-->
	<xsl:template match="/*">
		<xsl:message>
			<xsl:value-of select="concat (substring ($constHyp140, 1, $blockWidth), '&#xA;')"/>
			<xsl:text>
 *** ERROR: The input file cannot be processed due to mismatch in data layout. ***
                
</xsl:text>
			<xsl:value-of select="concat (substring ($constHyp140, 1, $blockWidth), '&#xA;')"/>
		</xsl:message>
	</xsl:template>
	<!--
		
	-->
	<xsl:template name="rec-count-lot-mismatch-xml">
		<xsl:param name="rows" />
		<xsl:param name="hits" select="0" />
		<xsl:choose>
			<xsl:when test="$rows [1]">
				<xsl:choose>
					<xsl:when test="$product [prd:material = $rows[1]/xls:Cell[1]/xls:Data]/prd:lotsize and not (0 = $rows[1]/xls:Cell[5]/xls:Data mod $product [prd:material = $rows[1]/xls:Cell[1]/xls:Data]/prd:lotsize)">
						<xsl:call-template name="rec-count-lot-mismatch-xml">
							<xsl:with-param name="rows" select="$rows [position () > 1]" />
							<xsl:with-param name="hits" select="$hits + 1" />
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="rec-count-lot-mismatch-xml">
							<xsl:with-param name="rows" select="$rows [position () > 1]" />
							<xsl:with-param name="hits" select="$hits" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$hits" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
		<!--
		
	-->
	<xsl:template name="rec-count-lot-mismatch-csv">
		<xsl:param name="objects" />
		<xsl:param name="hits" select="0" />
		<xsl:choose>
			<xsl:when test="$objects [1]">
				<xsl:variable name="material" select="$objects[1]/Property[@Name='Material']" />
				<xsl:choose>
					<xsl:when test="$product [prd:material = $material]/prd:lotsize and not (0 = $objects[1]/Property[@Name='Fakturierte Menge'] mod $product [prd:material = $material]/prd:lotsize)">
						<xsl:call-template name="rec-count-lot-mismatch-csv">
							<xsl:with-param name="objects" select="$objects [position () > 1]" />
							<xsl:with-param name="hits" select="$hits + 1" />
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="rec-count-lot-mismatch-csv">
							<xsl:with-param name="objects" select="$objects [position () > 1]" />
							<xsl:with-param name="hits" select="$hits" />
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$hits" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="xls:Row">
		<item xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<material>
				<xsl:value-of select="xls:Cell[1]/xls:Data" />
			</material>
			<kurztext>
				<xsl:value-of select="xls:Cell[2]/xls:Data" />
			</kurztext>
			<country>
				<xsl:value-of select="xls:Cell[7]/xls:Data" />
			</country>
			<units>
				<xsl:choose>
					<xsl:when test="$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize > 1">
						<xsl:attribute name="lot">
							<xsl:value-of select="$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize" />
						</xsl:attribute>
						<xsl:value-of select="ceiling (sum (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))/xls:Cell[5]/xls:Data) div $product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="sum (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))/xls:Cell[5]/xls:Data)" />
					</xsl:otherwise>
				</xsl:choose>
			</units>
		</item>
	</xsl:template>
	<!--
		**********************************************************************
			Das Template ...
			
		**********************************************************************
	-->
	<xsl:template match="xls:Table [xls:Row[1]/xls:Cell[1]/xls:Data = 'Material']">
		<!--
			
		-->
		<report xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<!--
				
			-->
			<xsl:if test="not ($product)">
				<xsl:message terminate="yes"> 
					<xsl:text>FATAL: Product master data missing. Aborting.</xsl:text>
				</xsl:message>
			</xsl:if>
			<!--
				
			-->
			<xsl:variable name="hits">
				<xsl:call-template name="rec-count-lot-mismatch-xml">
					<xsl:with-param name="rows" select="xls:Row [position () > 1][xls:Cell[7]/xls:Data = $SAP-DR-Recycling-Preprocess.Country]" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="$hits > 0">
				<messages xmlns="http://www.rothenberger.com/productcompliance/recycling/messages" number="{$hits}">
					<caption>
						<xsl:text>Items in documents where 'quantity' does not match the lot size:</xsl:text>
					</caption>
					<xsl:for-each select="xls:Row [position () > 1]">
						<xsl:if test="$product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize and not (0 = xls:Cell[5]/xls:Data mod $product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)">
							<message>
								<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1 - 1 - string-length (format-number (position(), '#.##0'))), format-number (position(), '#.##0'), ' ', xls:Cell [1]/xls:Data, substring ($constBlk140, 1, $indent2 - $indent1 - string-length (xls:Cell [1]/xls:Datal)), ' ', xls:Cell [2]/xls:Data, ' pcs.: ', xls:Cell [5]/xls:Data, ' lot size: ', $product [prd:material = current()/xls:Cell[1]/xls:Data]/prd:lotsize)" />
							</message>
						</xsl:if>
					</xsl:for-each>
				</messages>
			</xsl:if>
			<!--
				
			-->
			<xsl:apply-templates select="xls:Row [position () > 1][xls:Cell[7]/xls:Data = $SAP-DR-Recycling-Preprocess.Country][generate-id (.) = generate-id (key ('mat-xml', concat(xls:Cell[1]/xls:Data, '-', xls:Cell[7]/xls:Data))[1])]" />
		</report>
	</xsl:template>
	<!--
		**********************************************************************
			Das Template ...
			
		**********************************************************************
	-->
	<xsl:template match="xls:Worksheet[@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Table" />
	</xsl:template>
	<!--
		**********************************************************************
			Das Template ...
			
		**********************************************************************
	-->
	<xsl:template match="/xls:Workbook[xls:Worksheet/@ss:Name = 'Sheet1']">
		<xsl:apply-templates select="xls:Worksheet[@ss:Name = 'Sheet1']" />
	</xsl:template>
	<!--
		**********************************************************************
	-->
	<xsl:template match="Object [Property/@Name='Material' and Property/@Name='Fakturierte Menge' and Property/@Name='Empfangsland']" mode="wrapped">
		<xsl:variable name="material" select="Property[@Name='Material']" />
		<item xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<material>
				<xsl:value-of select="$material" />
			</material>
			<kurztext>
				<xsl:value-of select="Property[@Name='Materialkurztext']" />
			</kurztext>
			<country>
				<xsl:value-of select="Property[@Name='Empfangsland']" />
			</country>
			<units>
				<xsl:choose>
					<xsl:when test="$product [prd:material = $material]/prd:lotsize > 1">
						<xsl:attribute name="lot">
							<xsl:value-of select="$product [prd:material = $material]/prd:lotsize" />
						</xsl:attribute>
						<xsl:value-of select="ceiling (sum (key('mat-csv', concat (Property[@Name='Material'], '-', Property[@Name='Empfangsland']))/Property[@Name='Fakturierte Menge']) div $product [prd:material = $material]/prd:lotsize)" />
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="sum (key('mat-csv', concat (Property[@Name='Material'], '-', Property[@Name='Empfangsland']))/Property[@Name='Fakturierte Menge'])" />
					</xsl:otherwise>
				</xsl:choose>
				<!-- xsl:value-of select="count(key('mat-csv', concat (Property[@Name='Material'], '-', Property[@Name='Empfangsland'])))" /-->
			</units>
		</item>
	</xsl:template>
	<!--
		**********************************************************************
	-->
	<xsl:template match="Objects" mode="wrapped">
		<!--
			
		-->
		<report xmlns="http://www.rothenberger.com/productcompliance/recycling/commons">
			<!--
				
			-->
			<xsl:if test="not ($product)">
				<xsl:message terminate="yes"> 
					<xsl:text>FATAL: Product master data missing. Aborting.</xsl:text>
				</xsl:message>
			</xsl:if>
			<!--
				
			-->
			<xsl:variable name="hits">
				<xsl:call-template name="rec-count-lot-mismatch-csv">
					<xsl:with-param name="objects" select="Object[Property[@Name='Empfangsland'] = $SAP-DR-Recycling-Preprocess.Country]" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="$hits > 0">
				<messages xmlns="http://www.rothenberger.com/productcompliance/recycling/messages" number="{$hits}">
					<caption>
						<xsl:text>Items in documents where 'quantity' does not match the lot size:</xsl:text>
					</caption>
					<xsl:for-each select="Object [Property [@Name='Empfangsland'] = $SAP-DR-Recycling-Preprocess.Country]">
						<xsl:variable name="material" select="Property [@Name='Material']" />
						<xsl:variable name="kurztext" select="Property [@Name='Materialkurztext']" />
						<xsl:variable name="amount" select="Property [@Name='Fakturierte Menge']" />
						<xsl:if test="$product [prd:material = $material]/prd:lotsize and not (0 = $amount mod $product [prd:material = $material]/prd:lotsize)">
							<message>
								<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1 - 1 - string-length (format-number (position(), '#.##0'))), format-number (position(), '#.##0'), ' ', $material, substring ($constBlk140, 1, $indent2 - $indent1 - string-length ($material)), ' ', $kurztext, ' pcs.: ', $amount, ' lot size: ', $product [prd:material = $material]/prd:lotsize)" />
							</message>
						</xsl:if>
					</xsl:for-each>
				</messages>
			</xsl:if>
			<!--
				
			-->
			<!-- com:items -->
				<xsl:apply-templates select="Object[generate-id (.) = generate-id (key('mat-csv', concat (Property[@Name='Material'], '-', $SAP-DR-Recycling-Preprocess.Country))[1])]" mode="wrapped">
					<xsl:sort select="Property[@Name='Material']" />
				</xsl:apply-templates>
			<!-- /com:items -->
		</report>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="/Wrapped">
		<xsl:apply-templates select="Objects" mode="wrapped" />
	</xsl:template>
	<!--

	-->
	<xsl:template match="Property" />
	<!--

	-->
	<xsl:template match="Property [@Name='Fakturierte Menge']">
		<xsl:param name="amount" />
		<xsl:copy>
			<xsl:copy-of select="@Name" />
			<xsl:value-of select="$amount" />
		</xsl:copy>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Property [@Name='Material'] | Property[@Name='Materialkurztext'] | Property [@Name='Empfangsland']">
		<xsl:copy>
			<xsl:copy-of select="@Name" />
			<xsl:value-of select="normalize-space (.)" />
		</xsl:copy>
	</xsl:template>
	<!--

	-->
	<xsl:template match="Object">
		<xsl:variable name="amount">
			<xsl:choose>
				<xsl:when test="substring-after (Property [@Name='Fakturierte Menge'], $SAP-DR-Recycling-Preprocess.NumberDecimalSeparator)">
					<xsl:message>
						<xsl:text>WARNING: Sales volume should not be decimal formatted numbers.</xsl:text>
					</xsl:message>
					<xsl:value-of select="translate (substring-before (Property [@Name='Fakturierte Menge'], $SAP-DR-Recycling-Preprocess.NumberDecimalSeparator), concat ('0123456789', $SAP-DR-Recycling-Preprocess.NumberGroupSeparator), '0123456789')" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="translate (Property [@Name='Fakturierte Menge'], concat ('0123456789', $SAP-DR-Recycling-Preprocess.NumberGroupSeparator), '0123456789')" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:if test="not (number ($amount) = 0)">
			<xsl:copy>
				<xsl:apply-templates select="Property">
					<xsl:with-param name="amount" select="$amount" />
				</xsl:apply-templates>
			</xsl:copy>
		</xsl:if>
	</xsl:template>
	<!--
		**********************************************************************
	-->
	<xsl:template match="/Objects">
		<Wrapped>
			<xsl:copy>
				<xsl:apply-templates select="Object" />
			</xsl:copy>
		</Wrapped>
	</xsl:template>
	<!--
	
	-->
</xsl:stylesheet>
