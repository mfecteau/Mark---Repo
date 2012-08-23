<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<!-- XSLT parser for TSPD SampleMapper.xml file -->
	<xsl:template match="/">
		<html>
			<head>
				<title>TSPD Elements</title>
			</head>
			<body>
				<table width="100%" border="1" cellspacing="0" cellpadding="3">
					<tr align="center">
						<td colspan="2">
							<b>
								<font face="verdana" size="4">TSPD CONFIGURATION WORKSHEET</font>
							</b>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<font face="verdana" size="2">Client:</font>
						</td>
					</tr>
					<tr>
						<td>
							<font face="verdana" size="2">Client Contact: </font>
						</td>
						<td>
							<font face="verdana" size="2">Contact Phone:</font>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<font face="verdana" size="2">TSPD Configuration Specialist:</font>
						</td>
					</tr>
				</table><br/>&#160;
				<xsl:for-each select="*/ElementTab">
				<xsl:if test="@tabLabel != 'Document Elements' ">
					<br/>&#160;
					<table border="0" width="100%">
						<tr bgcolor="{@tabColor}">
							<td>
								<font face="verdana" size="4" color="#FFFFFF">
									<xsl:value-of select="@tabLabel"/>
									<xsl:value-of select="@tabColor"/>
								</font>
							</td>
						</tr>

						<table width="100%" border="1" frame="box" cellpadding="3" cellspacing="0"  >
							<tr bgcolor="black" >
							     
								<td width="25%">
									<font face="verdana" size="2" color="#FFFFFF">
										<b>TSPD Element</b>
									</font>
								</td>
								<td width="20%">
										<font face="verdana" size="2" color="#FFFFFF">
											<b>Configured Label</b>
											<br/>		     (leave blank if same)
										</font>	
								</td>
								<td width="10%">
										<font face="verdana" size="2" color="#FFFFFF">
											<b>Hidden</b>
										</font>	
								</td>
								<td width="10%">
										<font face="verdana" size="2" color="#FFFFFF">
											<b>Req'd</b>
										</font>	
								</td>
								<td width="10%">									
										<font face="verdana" size="2" color="#FFFFFF">
											<b>Default</b>
										</font>	
								</td>
								<td width="25%">									
										<font face="verdana" size="2" color="#FFFFFF">
											<b>ToolTip</b>
										</font>	
								</td>
							</tr>
							<xsl:for-each select="ElementBucket">
								<tr>
									<td>
										<font face="verdana" size="2">
										     <b>Bucket:
											<xsl:value-of select="@bucketLabel"/>
											</b>
										</font>
									</td>
									<td>
										<font face="verdana" size="2">&#160;
										<!--if different from what and where?xsl:value-of select="label"/-->
										</font>
									</td>
									<td>
										<font face="verdana" size="2">
											<xsl:choose >
												<xsl:when test="@hidden = 'true' or @hidden = 'True' ">  																<xsl:text>true</xsl:text>
												</xsl:when>
											</xsl:choose>
											&#160;
										</font>
									</td>										
									<td/>
									<td/>
									<td>
										<font face="verdana" size="2">
											<xsl:value-of select="@toolTip"/>
										</font>
									</td>
								</tr>
								<xsl:for-each select="ChooserEntry">
									<xsl:if test="count(ElementStatus) = 0">
										<tr>
											<td>
												<font face="verdana" size="2">																	<xsl:value-of select="@elementLabel"/>
												</font>
											</td>
											<td>
												<font face="verdana" size="2">&#160;
												</font>
											</td>
											<td>
												<font face="verdana" size="2">
													<xsl:choose >
														<xsl:when test="@hidden = 'true'  or @hidden = 'True' ">  																	<xsl:text>true</xsl:text>
														</xsl:when>
													</xsl:choose>
													&#160;
												</font>
											</td>										
											<td/>
											<td/>
											<td>
												<font face="verdana" size="2">
													<xsl:value-of select="@toolTip"/>
												</font>
											</td>
										</tr>
									</xsl:if>
									<xsl:if test="count(ElementStatus) = 1">
										<tr>
											<td>
												<font face="verdana" size="2">
													<xsl:value-of select="@elementLabel"/>
												</font>
											</td>
											<td>
												<font face="verdana" size="2">&#160;
												</font>
											</td>
											<td>
													<font face="verdana" size="2">
													<xsl:choose >
														<xsl:when test="@hidden = 'true' or @hidden = 'True' ">  																	<xsl:text>true</xsl:text>
														</xsl:when>
													</xsl:choose>
													&#160;
												</font>
												</td>
											<xsl:for-each select="ElementStatus">
												<td>
													<font face="verdana" size="2">
														<xsl:choose >
															<xsl:when test="@required = 'true' or @required = 'True' ">  																<xsl:text>true</xsl:text>
															</xsl:when>
														</xsl:choose>
														&#160;
													</font>
												</td>
												<td>
													<font face="verdana" size="2">
														<xsl:choose >
															<xsl:when test="@default = 'true' or @default = 'True' ">  																		<xsl:text>true</xsl:text>
															</xsl:when>
														</xsl:choose>
														&#160;
													</font>
												</td>
											</xsl:for-each>
											<td>
												<font face="verdana" size="2">
													<xsl:value-of select="@toolTip"/>
												</font>
											</td>
										</tr>
									</xsl:if>
									<xsl:if test="count(ElementStatus) > 1">
										<tr>
											<td>
												<font face="verdana" size="2">
													<xsl:value-of select="@elementLabel"/>
												</font>
											</td>
											<td>
												<font face="verdana" size="2">&#160;
												</font>
											</td>
											<td>
												<font face="verdana" size="2">
													<xsl:choose >
														<xsl:when test="@hidden = 'true'  or @hidden = 'True' ">  																	<xsl:text>truexyz</xsl:text>
														</xsl:when>
													</xsl:choose>
													&#160;
												</font>
											</td>										
											<td/>
											<td/>
											<td>
												<font face="verdana" size="2">
													<xsl:value-of select="@toolTip"/>
												</font>
											</td>
										</tr>
										<xsl:for-each select="ElementStatus">
											<tr>
												<td align="right">
													<font face="verdana" size="1">
														<xsl:text>Doc Type: </xsl:text>
														<xsl:value-of select="@docType"/>
													</font>
												</td>
												<td/>
												<td/>
												<td>
													<font face="verdana" size="2">
														<xsl:choose >
																<xsl:when test="@required = 'true' or @required = 'True' ">  																		<xsl:text>true</xsl:text>
															</xsl:when>
														</xsl:choose>
														&#160;																									</font>
												</td>
												<td>
													<font face="verdana" size="2">
														<xsl:choose >
															<xsl:when test="@default = 'true' or @default = 'True' ">  																		<xsl:text>true</xsl:text>
															</xsl:when>
														</xsl:choose>
														&#160;
													</font>
												</td>
												<td/>
											</tr>
										</xsl:for-each>
									</xsl:if>	
									<xsl:for-each select="Complex">
										<xsl:for-each select="ChooserEntry">
											<tr>
												<td align="right">
													<font face="verdana" size="2">
														<xsl:text>....</xsl:text>
														<xsl:value-of select="@elementLabel"/>
													</font>
												</td>
												<td>
													<font face="verdana" size="2">&#160;
													</font>
												</td>
												<td>
													<font face="verdana" size="2">
															<xsl:choose >
																<xsl:when test="@hidden = 'true' or @hidden = 'True' ">  																			<xsl:text>true</xsl:text>
																</xsl:when>
															</xsl:choose>
															&#160;
														</font>
													</td>
												<xsl:for-each select="ElementStatus">
													<td>
														<font face="verdana" size="2">
															<xsl:choose >
																<xsl:when test="@required = 'true' or @required = 'True' ">  																				<xsl:text>True</xsl:text>
																</xsl:when>
															</xsl:choose>
															&#160;
														</font>
													</td>
													<td>
														<font face="verdana" size="2">
															<xsl:choose >
																<xsl:when test="@default = 'true' or @default = 'True' ">  																			<xsl:text>true</xsl:text>
																</xsl:when>
															</xsl:choose>
														&#160;	
														</font>
													</td>
												</xsl:for-each>
												<xsl:if test="count(ElementStatus) = 0">							
													<td/>
													<td/>
												</xsl:if>
												<td>
													<font face="verdana" size="2">
														<xsl:value-of select="@toolTip"/>
													</font>
												</td>
											</tr>
										</xsl:for-each>
									</xsl:for-each>
								</xsl:for-each>
							</xsl:for-each>
						</table>
					</table>
</xsl:if>
				</xsl:for-each>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
