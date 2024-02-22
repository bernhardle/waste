<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:xs="http://www.w3.org/2001/XMLSchema"
	xmlns:fn="http://www.w3.org/2005/xpath-functions"
	xmlns:prd="http://www.rothenberger.com/productcompliance/recycling/productmasterdata"
	xmlns:com="http://www.rothenberger.com/productcompliance/recycling/commons"
	xmlns:dty="http://www.rothenberger.com/productcompliance/recycling/duties"
	exclude-result-prefixes="com dty fn prd xs">
	<!--

	-->
	<xsl:param name="Global.country" select="'DE'" />
	<xsl:param name="Global.batch" select="'WEEE'" />
	<!--

	-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--

	-->
	<xsl:key name="duties" match="dty:references/dty:entry" use="@SPKey" />
	<xsl:key name="munch" match="prd:product/prd:duties/prd:duty" use="concat (../../@SPKey, '-', @SPKey)" />
	<!--

	-->
	<xsl:template match="prd:product" mode="SCIP">
		<xsl:variable name="this" select="." />
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch]" />
			<item>
				<material>
					<xsl:value-of select="$this/prd:material" />
				</material>
				<kurztext>
					<xsl:value-of select="$this/prd:materialkurztext" />
				</kurztext>
				<batch>
					<xsl:value-of select="key ('duties', @SPKey)[1]/dty:batch" />
				</batch>
				<code>
					<xsl:value-of select="key ('duties', @SPKey)[1]/dty:code" />
				</code>
				<category>
					<xsl:value-of select="key ('duties', @SPKey)[1]/dty:label" />
				</category>
				<count>
					<xsl:value-of select="count ($current-group)" />
				</count>
			</item>
		</xsl:for-each>
	</xsl:template>
	<!--
		
	-->
	<xsl:template match="prd:product" mode="TVVV">
		<xsl:variable name="this" select="." />
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
		<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch]" />
		<item>
			<material>
				<xsl:value-of select="$this/prd:material" />
			</material>
			<kurztext>
				<xsl:value-of select="$this/prd:materialkurztext" />
			</kurztext>
			<batch>
				<xsl:value-of select="key ('duties', @SPKey)[1]/dty:batch" />
			</batch>
			<category>
				<xsl:value-of select="key ('duties', @SPKey)[1]/dty:label" />
			</category>
			<count>
				<xsl:value-of select="count ($current-group)" />
			</count>
			<alu>
				<xsl:value-of select="sum ($current-group/prd:data/VV_x002d_Alu)" />
			</alu>
			<steel>
				<xsl:value-of select="sum (current-group ()/prd:data/VV_x002d_Steel)" />
			</steel>
			<tinplate>
				<xsl:value-of select="sum ($current-group/prd:data/VV_x002d_Tinplate)" />
			</tinplate>
			<paper>
				<xsl:value-of select="sum ($current-group/prd:data/Paper)" />
			</paper>
			<plastic>
				<xsl:value-of select="sum ($current-group/prd:data/VV_x002d_Plastic)" />
			</plastic>
		</item>
		</xsl:for-each>
	</xsl:template>
	<!--

	-->
	<xsl:template match="prd:product" mode="WEEE-BATT">
		<xsl:variable name="this" select="." />
		<xsl:for-each select="$this/prd:duties/prd:duty [generate-id (.) = generate-id (key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch][1])]">
			<xsl:variable name="current-group" select="key ('munch', concat (../../@SPKey, '-', @SPKey))[key ('duties', @SPKey)[1]/com:countries/com:element = $Global.country][$Global.batch = key ('duties', @SPKey)[1]/dty:batch]" />
			<item>
				<material>
					<xsl:value-of select="$this/prd:material" />
				</material>
				<kurztext>
					<xsl:value-of select="$this/prd:materialkurztext" />
				</kurztext>
				<batch>
					<xsl:value-of select="key ('duties', @SPKey)[1]/dty:batch" />
				</batch>
				<category>
					<xsl:value-of select="key ('duties', @SPKey)[1]/dty:label" />
				</category>
				<count>
					<xsl:value-of select="count ($current-group)" />
				</count>
				<weight>
					<xsl:value-of select="sum ($current-group/prd:data/Weight)" />
				</weight>
			</item>
		</xsl:for-each>
	</xsl:template>
	<!--

	-->
	<xsl:template match="/prd:root">
		<list>
			<xsl:choose>
				<xsl:when test="$Global.batch = 'WEEE' or $Global.batch = 'BATT'">
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $Global.batch]/@SPKey]" mode="WEEE-BATT" />
				</xsl:when>
				<xsl:when test="$Global.batch = 'TVVV'">
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $Global.batch]/@SPKey]" mode="TVVV" />
				</xsl:when>
				<xsl:when test="$Global.batch = 'SCIP'">
					<xsl:apply-templates select="prd:product[prd:duties/prd:duty/@SPKey = current()/dty:references/dty:entry[com:countries/com:element = $Global.country and dty:batch = $Global.batch]/@SPKey]" mode="SCIP" />
				</xsl:when>
				<xsl:otherwise>
					<xsl:message>
						<xsl:text>[ERROR] Parameter Global.batch must be either WEEE, BATT or TVVV</xsl:text>
					</xsl:message>
				</xsl:otherwise>
			</xsl:choose>
		</list>
	</xsl:template>
	
</xsl:stylesheet>
