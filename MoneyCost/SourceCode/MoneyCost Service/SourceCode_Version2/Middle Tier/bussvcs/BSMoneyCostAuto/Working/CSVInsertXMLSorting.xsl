<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:template match="/">
	<CSV_INSERT_DATA_REQUEST>
	<CSV_INSERT_DATASet>
			<xsl:for-each select="CSV_INSERT_DATA_REQUEST/CSV_INSERT_DATASet/CSV_INSERT_DATA">
			<xsl:sort select="COL_POSITION" data-type="number" order="ascending"/>
				<CSV_INSERT_DATA>
					<COL_POSITION><xsl:value-of select="COL_POSITION"/></COL_POSITION>
					<INTEREST_RATE><xsl:value-of select="INTEREST_RATE"/></INTEREST_RATE>
				</CSV_INSERT_DATA>
			</xsl:for-each>
	</CSV_INSERT_DATASet>
	</CSV_INSERT_DATA_REQUEST>
	</xsl:template>
</xsl:stylesheet>