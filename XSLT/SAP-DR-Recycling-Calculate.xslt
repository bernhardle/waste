<?xml version="1.0" encoding="UTF-8"?>
<!--
	(c) Bernhard Schupp, Frankfurt (2021-2022)
		
	Revision:
		2021-02-21:	Erstellt.
		2021-03-01:	Tin & Iron separated.
		2021-05-04:	Sorting by field dty:rank and label including dty:rank as Line # added.
		2022-06-05:	Support for 'Product-in-Lots' added.
		2022-09-29:	Second master data integrity check only in verbose mode.
		2022-10-02:	Major change ready to support products in lots.
-->
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:ns1="http://www.w3.org/2005/Atom" 
	xmlns:xlink="http://www.w3.org/1999/xlink" 
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata" 
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons" 
	xmlns:msg="http://www.rothenberger.com/productcompliance/recycling/messages" 
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty prd">
	<!--
		
	-->
	<xsl:param name="debug" select="0" />
	<xsl:param name="verbose" select="0" />
	<xsl:param name="SAP-DR-Recycling-Converter.Country" select="'DE'" />
	<xsl:param name="SAP-DR-Recycling-Converter.MasterData" />
	<!--

	-->
	<xsl:variable name="indent1" select="10" />
	<xsl:variable name="indent2" select="22" />
	<xsl:variable name="indent3" select="45" />
	<xsl:variable name="indent4" select="70" />
	<!--
		
	-->
	<xsl:variable name="constBlk140" select="'                                                                                                                                            '" />
	<xsl:variable name="constHyp140" select="'--------------------------------------------------------------------------------------------------------------------------------------------'" />
	<xsl:variable name="constDot140" select="'............................................................................................................................................'" />
	<xsl:variable name="blockWidth" select="95" />
	
	<xsl:variable name="master" select="document ($SAP-DR-Recycling-Converter.MasterData)/prd:root" />
	<xsl:variable name="batch" select="$master/@batch" />
	<xsl:variable name="filter" select="$master/dty:references/dty:entry [com:countries/com:element = $SAP-DR-Recycling-Converter.Country]/@SPKey" />
	<xsl:variable name="duties" select="$master/prd:product/prd:duties/prd:duty[@SPKey = $filter]" />
	<xsl:variable name="products" select="$master/prd:product" />
	
	<xsl:key name="mat" use="com:material" match="com:item" />
	<xsl:key name="cat" use="@SPKey" match="prd:product/prd:duties/prd:duty" />
	<xsl:key name="prd" use="prd:product/prd:duties/prd:duty/@SPKey" match="prd:product" />
	<!--
		
	-->
	<xsl:decimal-format decimal-separator="," grouping-separator="." NaN="*" />
	<!--
		
	-->
	<xsl:output method="text" encoding="UTF-8" />
	<!--
		
		**********************************************************************
		
	-->
	<xsl:template name="format">
		<xsl:param name="hdr" />
		<xsl:param name="lab" />
		<xsl:param name="pcs" />
		<xsl:param name="wgt" />
		<xsl:variable name="label" select="substring($lab, 1, 20)" />
		<xsl:variable name="pieces" select="format-number($pcs, '#.##0')" />
		<xsl:variable name="weight" select="format-number($wgt, '#.##0,000')" />
		<xsl:if test="$hdr">
			<xsl:value-of select="concat(substring ($constBlk140, 1, $indent1), $hdr, ':', '&#xA;&#xA;')" />
		</xsl:if>
		<xsl:value-of select="concat(substring($constBlk140, 1, $indent3), $lab, ' ', substring ($constDot140, 1, $indent4 - $indent3 - string-length ($lab) - string-length ($pieces)), ' ', $pieces, ' pcs.', substring ($constBlk140, 1, 12 - string-length ($weight)), $weight, ' kg', '&#xA;&#xA;')" />
	</xsl:template>
	<!--
		
		**********************************************************************
		
	-->
	<xsl:template name="batt-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="scr" select="0" />
		<xsl:param name="pcs" select="0" />
		<xsl:param name="ces" select="0" />
		<!--
			
		-->
		<xsl:variable name="pos" select="$duties [1]" />
		<!--

		-->
		<xsl:choose>
			<xsl:when test="$pos">
				<xsl:variable name="rows" select="$table [com:material = $pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="batt-rec">
					<xsl:with-param name="header" select="$header" />
					<xsl:with-param name="duties" select="$duties [position() > 1]" />
					<xsl:with-param name="table" select="$table" />
					<xsl:with-param name="inv" select="$inv" />
					<xsl:with-param name="cnt" select="$cnt + sum($rows/com:units)" />
					<xsl:with-param name="scr" select="$scr + sum($rows/com:units) * $pos/prd:data/Weight" />
					<xsl:with-param name="pcs" select="$pcs + sum($rows/com:units)" />	
					<!-- xsl:with-param name="pcs" select="$pcs + sum($rows/com:units) * $pos/prd:data/Pieces" /-->
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$verbose or $cnt > 0 or $scr > 0">
					<xsl:call-template name="format">
						<xsl:with-param name="hdr" select="$header" />
						<xsl:with-param name="lab" select="'Total:'" />
						<xsl:with-param name="pcs" select="$pcs" />
						<xsl:with-param name="wgt" select="$scr" />
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
		
		**********************************************************************
		
	-->
	<xsl:template name="weee-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="scr" select="0" />
		<!--
			
		-->
		<xsl:variable name="pos" select="$duties [1]" />
		<!--

		-->
		<xsl:choose>
			<xsl:when test="$pos">
				<xsl:variable name="rows" select="$table [com:material = $pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="weee-rec">
					<xsl:with-param name="header" select="$header" />
					<xsl:with-param name="duties" select="$duties [position() > 1]" />
					<xsl:with-param name="table" select="$table" />
					<xsl:with-param name="inv" select="$inv" />
					<xsl:with-param name="cnt" select="$cnt + sum($rows/com:units)" />
					<xsl:with-param name="scr" select="$scr + sum($rows/com:units) * $pos/prd:data/Weight" />	
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$verbose or $cnt > 0 or $scr > 0">
					<xsl:call-template name="format">
						<xsl:with-param name="hdr" select="$header" />
						<xsl:with-param name="lab" select="'Total:'" />
						<xsl:with-param name="pcs" select="$cnt" />
						<xsl:with-param name="wgt" select="$scr" />
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
		
		**********************************************************************
		
	-->
	<xsl:template name="tvvv-rec">
		<xsl:param name="header" select="'Hier sollte der Titel stehen.'" />
		<xsl:param name="table" />
		<xsl:param name="duties" />
		<xsl:param name="inv" select="0" />
		<xsl:param name="cnt" select="0" />
		<xsl:param name="alu" select="0" />
		<xsl:param name="alu-cnt" select="0" />
		<xsl:param name="iron" select="0" />
		<xsl:param name="iron-cnt" select="0" />
		<xsl:param name="ppk" select="0" />
		<xsl:param name="ppk-cnt" select="0" />
		<xsl:param name="plast" select="0" />
		<xsl:param name="plast-cnt" select="0" />
		<xsl:param name="tin" select="0" />
		<xsl:param name="tin-cnt" select="0" />
		<!--
			
		-->
		<xsl:variable name="pos" select="$duties [1]" />
		<!--

		-->
		<xsl:choose>
			<xsl:when test="$pos">
				<xsl:variable name="rows" select="$table [com:material = $pos/parent::prd:duties/parent::prd:product/prd:material]" />
				<xsl:call-template name="tvvv-rec">
					<xsl:with-param name="header" select="$header" />
					<xsl:with-param name="duties" select="$duties [position() > 1]" />
					<xsl:with-param name="table" select="$table" />
					<xsl:with-param name="inv" select="$inv" />
					<xsl:with-param name="cnt" select="$cnt + sum($rows/com:units)" />
					<xsl:with-param name="alu" select="$alu + sum($rows/com:units) * $pos/prd:data/VV_x002d_Alu" />
					<xsl:with-param name="iron" select="$iron + sum($rows/com:units) * $pos/prd:data/VV_x002d_Steel" />
					<xsl:with-param name="ppk" select="$ppk + sum($rows/com:units) * $pos/prd:data/Paper" />
					<xsl:with-param name="plast" select="$plast + sum($rows/com:units) * $pos/prd:data/VV_x002d_Plastic" />
					<xsl:with-param name="tin" select="$tin + sum($rows/com:units) * $pos/prd:data/VV_x002d_Tinplate" />
					<xsl:with-param name="alu-cnt">
						<xsl:choose>
							<xsl:when test="$pos/prd:data/VV_x002d_Alu > 0">
								<xsl:value-of select="$alu-cnt + sum($rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$alu-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="iron-cnt">
						<xsl:choose>
							<xsl:when test="$pos/prd:data/VV_x002d_Steel > 0">
								<xsl:value-of select="$iron-cnt + sum($rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$iron-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="ppk-cnt">
						<xsl:choose>
							<xsl:when test="$pos/prd:data/Paper > 0">
								<xsl:value-of select="$ppk-cnt + sum($rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$ppk-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="plast-cnt">
						<xsl:choose>
							<xsl:when test="$pos/prd:data/VV_x002d_Plastic > 0">
								<xsl:value-of select="$plast-cnt + sum($rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$plast-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
					<xsl:with-param name="tin-cnt">
						<xsl:choose>
							<xsl:when test="$pos/prd:data/VV_x002d_Tinplate > 0">
								<xsl:value-of select="$tin-cnt + sum($rows/com:units)" />
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$tin-cnt" />
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<!--
				
			-->
			<xsl:otherwise>
				<xsl:call-template name="format">
					<xsl:with-param name="hdr" select="$header" />
					<xsl:with-param name="lab" select="'Paper'" />
					<xsl:with-param name="wgt" select="$ppk * 0.001" />
					<xsl:with-param name="pcs" select="$ppk-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Tinplate'" />
					<xsl:with-param name="wgt" select="$tin * 0.001" />
					<xsl:with-param name="pcs" select="$tin-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Alu'" />
					<xsl:with-param name="wgt" select="$alu * 0.001" />
					<xsl:with-param name="pcs" select="$alu-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Iron'" />
					<xsl:with-param name="wgt" select="$iron * 0.001" />
					<xsl:with-param name="pcs" select="$iron-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Plastics'" />
					<xsl:with-param name="wgt" select="$plast* 0.001" />
					<xsl:with-param name="pcs" select="$plast-cnt" />
				</xsl:call-template>
				<xsl:call-template name="format">
					<xsl:with-param name="lab" select="'Total'" />
					<xsl:with-param name="wgt" select="($alu + $iron + $tin + $ppk + $plast) * 0.001" />
					<xsl:with-param name="pcs" select="$cnt" />
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--
		
		**********************************************************************
		
	-->
	<xsl:template match="com:item" mode="check">
		<xsl:variable name="total" select="sum (key('mat', com:material)[com:country = $SAP-DR-Recycling-Converter.Country]/com:units)"/>
		<xsl:variable name="col1w" select="$indent1"/>
		<xsl:variable name="col2w" select="$indent2 - $indent1"/>
		<xsl:variable name="col3w" select="40"/>
		<xsl:variable name="col4w" select="8"/>
		<xsl:variable name="col5w" select="6"/>
		<xsl:variable name="col6w" select="8"/>
		<xsl:value-of select="concat (substring ($constBlk140, 1, $col1w), com:material, substring ($constBlk140, 1, $col2w - string-length (com:material)), substring (com:kurztext, 1, $col3w))"/>
		<xsl:value-of select="concat (substring ($constBlk140, 1, $col5w + $col4w + $col3w - string-length (substring (com:kurztext, 1, $col3w))), substring ($constBlk140, 1, $col6w - string-length (format-number ($total, '#.##0'))), format-number ($total, '#.##0'), ' pcs.&#xA;')" />
	</xsl:template>
	<!--

	-->
	<xsl:template match="msg:messages">
		<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1), '********* ', msg:caption,'  *********', '&#xA;&#xA;')"/>
		<xsl:for-each select="child::msg:message">
			<xsl:value-of select="concat(. ,'&#xA;')" />
		</xsl:for-each>
		<xsl:value-of select="concat ('&#xA;', substring($constHyp140, 1, $blockWidth), '&#xA;&#xA;')" />
	</xsl:template>
	<!--
		**********************************************************************
			Das Template listet alle Materialnummern auf, zu denen im betreffenden
			Verkaufsland keine Duties in den Stammdaten hinterlegt sind. 
		**********************************************************************
	-->
	<xsl:template match="com:report" mode="check">
		<xsl:variable name="nom" select="com:item [com:country = $SAP-DR-Recycling-Converter.Country][not(com:material = $duties/parent::prd:duties/parent::prd:product/prd:material)][generate-id (key('mat', com:material)[com:country = $SAP-DR-Recycling-Converter.Country][1]) = generate-id (.)]" />
		<xsl:if test="$nom">
			<xsl:if test="$nom [not(com:material = $master/prd:product/prd:material)]">
				<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1), '********* Produkte in Belegen, zu denen die Produkt-Stammdaten fehlen *********', '&#xA;&#xA;')"/>
				<xsl:apply-templates select="$nom [not(com:material = $master/prd:product/prd:material)]" mode="check">
					<xsl:sort select="com:material" />
				</xsl:apply-templates>
				<xsl:value-of select="'&#xA;'" />
			</xsl:if>
			<xsl:if test="$verbose">
				<xsl:if test="$nom [com:material = $master/prd:product/prd:material]">
					<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1), '********* Produkte in Belegen, ohne Duties im Batch ', $batch, ' und Land ', $SAP-DR-Recycling-Converter.Country ,'  *********', '&#xA;&#xA;')"/>
					<xsl:apply-templates select="$nom [com:material = $master/prd:product/prd:material]" mode="check">
						<xsl:sort select="com:material" />
					</xsl:apply-templates>
					<xsl:value-of select="'&#xA;'" />
				</xsl:if>
			</xsl:if>
			<xsl:if test="not(msg:messages)">
				<xsl:value-of select="concat (substring($constHyp140, 1, $blockWidth), '&#xA;&#xA;')" />
			</xsl:if>
		</xsl:if>
		<xsl:apply-templates select="msg:messages" />
	</xsl:template>
	<!--
		**********************************************************************
			Das Template summiert alle 'duty' Knoten in den Produkten zu der betreffenden
			Kategorie mit den Daten aus der Verkaufsliste in einer rekursiven Verarbeitung.
		**********************************************************************
	-->
	<xsl:template match="com:report" mode="run">
		<xsl:variable name="here" select="current()" />
		<xsl:for-each select="$duties [generate-id (key('cat', @SPKey)[1]) = generate-id (.)]">
			<xsl:sort select="$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank" data-type="number" order="ascending" />
			<xsl:variable name="filter" select="key('cat',@SPKey)" />
			<xsl:variable name="rowx" select="$here/com:item [com:country = $SAP-DR-Recycling-Converter.Country][com:material = $filter/parent::prd:duties/parent::prd:product/prd:material]" />
			<xsl:choose>
				<xsl:when test="$batch = 'BATT'">
					<xsl:call-template name="batt-rec">
						<xsl:with-param name="header" select="concat ('[Line ', $master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank, '] ', $master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label)" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="$rowx" />
						<xsl:with-param name="inv" select="count($rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:when test="$batch = 'TVVV'">
					<xsl:call-template name="tvvv-rec">
						<xsl:with-param name="header" select="$master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="$rowx" />
						<xsl:with-param name="inv" select="count($rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:when test="$batch = 'WEEE'">
					<xsl:call-template name="weee-rec">
						<xsl:with-param name="header" select="concat ('[Line ', $master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:rank, '] ', $master/dty:references/dty:entry[@SPKey = current()/@SPKey]/dty:label)" />
						<xsl:with-param name="duties" select="key('cat',@SPKey)" />
						<xsl:with-param name="table" select="$rowx" />
						<xsl:with-param name="inv" select="count($rowx)" />
					</xsl:call-template>			
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="concat ('&#xA;', substring ($constHyp140, 1, $blockWidth), '&#xA;', '&#xA;')"/>
					<xsl:value-of select="concat (' *** FEHLER: Der Batch &quot;', $batch, '&quot; ist nicht vorgesehen. ***', '&#xA;')"/>
					<xsl:value-of select="concat ('&#xA;', substring ($constHyp140, 1, $blockWidth), '&#xA;', '&#xA;')"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		
	</xsl:template>
	<!--
		**********************************************************************
			Verteiler Template
		**********************************************************************
	-->
	<xsl:template match="/com:report">
		<xsl:value-of select="concat ('&#xA;', substring($constHyp140, 1, $blockWidth), '&#xA;&#xA;')" />
		<xsl:value-of select="concat (substring ($constBlk140, 1, $indent1), $SAP-DR-Recycling-Converter.Country, '&#xA;&#xA;')" />
		<xsl:apply-templates select="." mode="check" />
		<xsl:apply-templates select="." mode="run" />
		<xsl:value-of select="concat ('&#xA;', substring($constHyp140, 1, $blockWidth), '&#xA;')" />
	</xsl:template>
	<!--
		**********************************************************************
	-->
</xsl:stylesheet>
