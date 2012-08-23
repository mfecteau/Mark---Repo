<?xml version='1.0'?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<!-- XSLT parser for TSPD Mapper.xml file -->
  <xsl:template match="/">
      <html>
         <head>
	 <title>TSPD Elements</title>   
	 </head>
	 <body>
	 <xsl:for-each select="*/ElementTab">
	 <table border="0" width="90%">
	 <tr bgcolor="{@tabColor}">
	 <td><font face="verdana" size="4" color="#FFFFFF">
	 <xsl:value-of select="@tabLabel"/><xsl:value-of select="@tabColor"/></font>
	 </td>
	 </tr>
	 </table>
    </xsl:for-each>
	</body>
      </html>
  </xsl:template>
</xsl:stylesheet>