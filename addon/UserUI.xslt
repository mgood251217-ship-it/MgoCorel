<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:frmwrk="Corel Framework Data">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:applicationInfo userConfiguration="true"/>
    <frmwrk:compositeNode xPath="/uiConfig/commandBars/commandBarData[@guid='a84d5e31-b1d6-4f8e-9c27-6274a0e11f35']"/>
    <frmwrk:compositeNode xPath="/uiConfig/frame"/>
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/states/state[@name='frame']//dockHost[@orientation='horizontal' and @dock='top'][1]">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(toolbar[@guidRef='a84d5e31-b1d6-4f8e-9c27-6274a0e11f35'])">
        <toolbar guidRef="a84d5e31-b1d6-4f8e-9c27-6274a0e11f35" userVisibility="true" dock="top"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
