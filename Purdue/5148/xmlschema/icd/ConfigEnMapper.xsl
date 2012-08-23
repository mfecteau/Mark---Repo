<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<!-- XSLT parser for TSPD Mapper.xml file -->
	<xsl:template match="/">
		<html>
			<head>
				<title>TSPD PICKLISTS</title>
			</head>
			<body>
				<table width="100%" cellspacing="0" frame="box" >
					<tr align="center">
						<td colspan="2">
							<b>
								<font face="verdana" size="4">TSPD PICKLISTS CONFIGURATION</font>
							</b>
						</td>
					</tr>
					<tr>
						<td colspan="1" width="25%">
							<font face="verdana" size="2"><b>Requested by:</b></font>
						</td>
						<td/>
					</tr>
					<tr>
						<td colspan="1" width="25%">
							<font face="verdana" size="2"><b>Requested date:</b></font>
						</td>
						<td/>
					</tr>
					<tr>
						<td colspan="1" width="25%">
							<font face="verdana" size="2"><b>For client:</b></font>
						</td>
						<td/>
					</tr>
				</table><br/>&#160;
				<table width="100%"  frame="box" cellspacing="0" >
					<tr bgcolor="black" align="center" >
						<td width="34%">
							<font face="verdana" size="2" color="#FFFFFF">
								<b>List / Internal Value</b>
							</font>
						</td>
						<td width="33%">
								<font face="verdana" size="2" color="#FFFFFF">
									<b>Default User Choice</b>
								</font>	
						</td>
						<td width="33%">
								<font face="verdana" size="2" color="#FFFFFF">
									<b>New Client Value</b>
								</font>	
						</td>
					</tr>
					
					
					<xsl:for-each select="*/EnumType">
						<tr charoff="3">
							<td  colspan="3">&#160;&#160;
								<font face="verdana" size="2">
									<b><xsl:value-of select="@enumName"/></b>
								</font>
							</td>
						</tr>
						<!--EnumPair systemName="I" userLabel="1"/-->
						<xsl:for-each select="EnumPair">
							<tr>
								<td>&#160;&#160;
									<font face="verdana" size="2">
										<xsl:value-of select="@systemName"/>
									</font>
								</td>
								<td>&#160;&#160;
									<font face="verdana" size="2">
										<xsl:value-of select="@userLabel"/>
									</font>
								</td>
								<td>&#160;</td>
							</tr>
						</xsl:for-each>
						<tr>
							<td colspan="3">&#160;</td>
						</tr>
						<tr>
							<td colspan="3">&#160;</td>
						</tr>
						<tr>
							<td colspan="3">&#160;</td>
						</tr>
					</xsl:for-each>
				</table>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
