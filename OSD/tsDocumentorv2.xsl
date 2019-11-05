<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:template match="/">
<HTML>
<HEAD>
<!-- Version 1.1
Changes:
    - Formats orphan steps with no Group
    - Added recursive template for Nested conditions for both groups and steps (tested to 3 levels)
    - Properly convert Not to Nor for conditions
' * =======================================================================================
'*
'* Disclaimer
'*
'* This script is not supported under any Microsoft standard support program or service. This 
'* script is provided AS IS without warranty of any kind. Microsoft further disclaims all 
'* implied warranties including, without limitation, any implied warranties of merchantability 
'* or of fitness for a particular purpose. The entire risk arising out of the use or performance 
'* of this script remains with you. In no event shall Microsoft, its authors, or anyone else 
'* involved in the creation, production, or delivery of this script be liable for any damages 
'* whatsoever (including, without limitation, damages for loss of business profits, business 
'* interruption, loss of business information, or other pecuniary loss) arising out of the use 
'* of or inability to use this script, even if Microsoft has been advised of the possibility 
'* of such damages.
'*
'*=======================================================================================
-->
    <STYLE TYPE="text/css">
      TD.group { background-color:teal;color:white }
      TD.step { background-color:beige }
      TD.header { background-color:black;color:white }
    </STYLE>
</HEAD>

<BODY>
<TABLE border='0' cellpadding='2' cellspacing ='0' style='font: 9px arial;border-width:0px;border-spacing:0px;border-style:none' width="100%" >

<!-- Header row for Task Sequence -->
<TR >
		<TD Class='header'><B>Group</B></TD>
		<!-- <TD Class='header'><B>Description</B></TD>
		<TD Class='header'></TD> -->
		<TD Class='header'><B>Conditions</B></TD>
		<TD Class='header'></TD>
		<xsl:for-each select="descendant::group">
            <TD Class='header'></TD>
    </xsl:for-each>
</TR> 
    
<!-- Parse Each Group or Orphan steps that are part of no groups -->
<xsl:for-each select="SmsTaskSequencePackage/SequenceData/sequence//group | sequence//group | SmsTaskSequencePackage/SequenceData/sequence | sequence">
	<TR>
        <!-- Indent for parent groups, i.e. add cells for the number of group ancestors -->
        <xsl:for-each select="ancestor::group">
            <TD></TD>
        </xsl:for-each>
        
        <!-- Add the group name and description -->
        <TD Class='group'><xsl:value-of select="@name"/></TD>
        <!-- <TD Class='group'><xsl:value-of select="@description"/></TD> -->
        <TD Class='group'>
            <!-- Output the continue on error flag -->
            <xsl:if test="@continueOnError='true'">
                    continue on error<BR />
            </xsl:if>
                
            <!-- Conditions with preceding operators -->
            <xsl:for-each select="condition/operator">
                <xsl:call-template name="FormatOperators">
                      <xsl:with-param name="RootNode" select="." />
                </xsl:call-template>
            </xsl:for-each>
             
            <!-- Conditions with no preceding operators -->
            <xsl:if test="not(./condition//operator)">
                <xsl:for-each select="./condition//expression">
                      <xsl:value-of select=".//variable[@name='Query']"/>
                      <xsl:value-of select=".//variable[@name='Variable']"/><BR />
                      <xsl:value-of select=".//variable[@name='Operator']"/><BR />
                      <xsl:value-of select=".//variable[@name='Value']"/><BR />
                </xsl:for-each>    
            </xsl:if>
        </TD>
         
		<!-- Trailing cells for child groups, i.e. add cells for the number of group decendants -->
		<!-- <TD Class='group'></TD> -->
		<TD Class='group'></TD>
		<TD Class='group'></TD>
		<xsl:for-each select="descendant::group">
         <TD Class='group'></TD>
    </xsl:for-each>
	</TR>
	
    <!-- Parse Each Task(Step) Display Header -->
    <TR >
        <!-- Indent for parent groups, i.e. add cells for the number of group ancestors -->    
        <xsl:for-each select="ancestor::group">
            <TD></TD>
        </xsl:for-each>
        <!-- Display header only if a child step exists -->
        <xsl:if test="child::step">
            <TD><B>Task</B></TD>
            <TD><B>Conditions</B></TD>
            <!-- <TD><B>Description</B></TD> -->
            <!-- <TD><B>Package</B></TD> -->
            <!-- <TD><B>Action</B></TD> -->
            <!-- <TD><B>Variables</B></TD> -->
        </xsl:if>
    </TR>
    
	<!-- Parse Each Task(Step) that is not disabled -->
	<xsl:for-each select="step[not(@disable='true')]">
        <TR >
        <!-- Indent for parent groups of parent, i.e. add cells for the number of group ancestors of the parent group -->    
        <xsl:for-each select="parent::group/ancestor::group">
            <TD ></TD>
        </xsl:for-each>

        <!-- Output Name and Description of Step -->
        <TD Class='step'><xsl:value-of select="@name"/></TD>
        <!-- <TD Class='step'><xsl:value-of select="@description"/></TD> -->

        <!-- Conditions for steps -->
        <TD Class='step'>
            <xsl:if test="@continueOnError='true'">
                continue on error<BR />
            </xsl:if>

            <!-- Conditions with preceding operators -->
            <xsl:for-each select="condition/operator">
                <xsl:call-template name="FormatOperators">
                      <xsl:with-param name="RootNode" select="." />
                </xsl:call-template>
            </xsl:for-each>
             
            <!-- Conditions with no preceding operators -->
            <xsl:if test="not(./condition/operator)">
                <xsl:for-each select="./condition//expression">
                      <xsl:value-of select=".//variable[@name='Query']"/>
                      <xsl:value-of select=".//variable[@name='Variable']"/><BR />
                      <xsl:value-of select=".//variable[@name='Operator']"/><BR />
                      <xsl:value-of select=".//variable[@name='Value']"/><BR />
                </xsl:for-each>    
            </xsl:if>
        </TD>
        <!-- Output the PackageID, DriverPackageID or CustomSettingPackageID or OSPackageIDs for the step -->
        <!-- <TD Class='step'>
            <xsl:for-each select=".//variable[@name='PackageID' or @name='OSDApplyDriverDriverPackageID' or @name='ConfigFilePackage' or @name='ImagePackageID']">
                <xsl:value-of select="."/><BR />
            </xsl:for-each>    
        </TD> -->
        <!-- Output the step action -->
        <!-- <TD Class='step'><xsl:value-of select="action"/></TD> -->
        <!-- <TD Class='step'> -->
            <!-- Output non-commandline variables (commandline should be in previous action field) -->
            <!-- <xsl:for-each select="defaultVarList/variable[not(@name='CommandLine')]"> -->
                <!-- <xsl:value-of select="./@name"/>:<xsl:value-of select="." /><BR/> -->
            <!-- </xsl:for-each>     -->
        <!-- </TD> -->
        </TR>
        
    </xsl:for-each> <!-- Step -->
        
</xsl:for-each> <!-- Group -->

</TABLE>
</BODY>
</HTML>
</xsl:template>

<xsl:template name="FormatOperators">
<xsl:param name="RootNode" />
     <xsl:if test="not($RootNode/preceding-sibling::expression | $RootNode/preceding-sibling::operator)"><xsl:text> </xsl:text>(<xsl:text> </xsl:text></xsl:if>
        <xsl:for-each select="$RootNode/operator">
             <xsl:call-template name="FormatOperators">
                  <xsl:with-param name="RootNode" select="." />
             </xsl:call-template>
        </xsl:for-each>
        <xsl:for-each select="$RootNode/expression">
             <xsl:if test="following-sibling::operator">
                <!-- Convert "NOT" Operator to "NOR" -->
                <xsl:if test="parent::operator/@type='not'">nor</xsl:if>
                <xsl:if test="not(parent::operator/@type='not')"><xsl:value-of select="parent::operator/@type"/></xsl:if>
             </xsl:if>
             <xsl:if test="not(preceding-sibling::expression | preceding-sibling::operator)"><xsl:text> </xsl:text>(</xsl:if>
             <xsl:value-of select=".//variable[@name='Query']"/><xsl:text> </xsl:text>
             <xsl:value-of select=".//variable[@name='Variable']"/><xsl:text> </xsl:text>
             <xsl:value-of select=".//variable[@name='Operator']"/><xsl:text> </xsl:text>
             <xsl:value-of select=".//variable[@name='Value']"/><xsl:text> </xsl:text>
             <xsl:if test="following-sibling::expression">
                <!-- Convert "NOT" Operator to "NOR" -->
                <xsl:if test="parent::operator/@type='not'">nor</xsl:if>
                <xsl:if test="not(parent::operator/@type='not')"><xsl:value-of select="parent::operator/@type"/></xsl:if>
             </xsl:if>
             <xsl:if test="not(following-sibling::expression | following-sibling::operator)"><xsl:text> </xsl:text>)<xsl:text> </xsl:text></xsl:if>
        </xsl:for-each>
    <xsl:if test="$RootNode/following-sibling::expression | $RootNode/following-sibling::operator">
            <!-- Convert "NOT" Operator to "NOR" -->
            <xsl:if test="parent::operator/@type='not'">nor</xsl:if>
            <xsl:if test="not(parent::operator/@type='not')"><xsl:value-of select="parent::operator/@type"/></xsl:if>
    </xsl:if><xsl:text> </xsl:text>
    <xsl:if test="not($RootNode/following-sibling::expression | $RootNode/following-sibling::operator)"><xsl:text> </xsl:text>)<xsl:text> </xsl:text></xsl:if>
</xsl:template>

</xsl:stylesheet>
