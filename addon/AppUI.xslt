<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:frmwrk="Corel Framework Data">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:applicationInfo userConfiguration="true"/>
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/items">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>

      <itemData guid="36f57221-bc31-4ddb-8c66-5c86b1f999b6"
                icon="guid://36f57221-bc31-4ddb-8c66-5c86b1f999b6"
                type="flyout"
                flyoutBarRef="812bc65f-0465-470e-97b2-6139a95a0001"
                userCreated="true"
                noBmpOnMenu="true"
                userCaption="Detail"
                userToolTip="Detail dan Pengaturan"/>

      <itemData guid="c64401b3-ea8a-4295-bbcb-d8c407dba836"
                icon="guid://c64401b3-ea8a-4295-bbcb-d8c407dba836"
                type="flyout"
                flyoutBarRef="812bc65f-0465-470e-97b2-6139a95a0002"
                userCreated="true"
                noBmpOnMenu="true"
                userCaption="Spanduk"
                userToolTip="Fungsi Spanduk"/>

      <itemData guid="15f0179a-5594-4ec8-b3fd-69ee17eedcfa"
                icon="guid://15f0179a-5594-4ec8-b3fd-69ee17eedcfa"
                type="flyout"
                flyoutBarRef="812bc65f-0465-470e-97b2-6139a95a0003"
                userCreated="true"
                noBmpOnMenu="true"
                userCaption="Layout"
                userToolTip="Fungsi Layout"/>

      <itemData guid="7f6d3bd7-7144-47a6-aaae-659ba4beb89d"
                icon="guid://7f6d3bd7-7144-47a6-aaae-659ba4beb89d"
                type="flyout"
                flyoutBarRef="812bc65f-0465-470e-97b2-6139a95a0004"
                userCreated="true"
                noBmpOnMenu="true"
                userCaption="Produksi"
                userToolTip="Fungsi Produksi"/>

      <itemData guid="4d9a55ff-0ad7-4268-8a7e-dd8fa602cc91"
                icon="guid://4d9a55ff-0ad7-4268-8a7e-dd8fa602cc91"
                type="flyout"
                flyoutBarRef="812bc65f-0465-470e-97b2-6139a95a0005"
                userCreated="true"
                noBmpOnMenu="true"
                userCaption="Job"
                userToolTip="Fungsi Job"/>

      <itemData guid="38533386-1283-49fa-9e35-124f6fab8744"
                icon="guid://38533386-1283-49fa-9e35-124f6fab8744"
                dynamicCommand="MgoCorel.modMain.Plugin_ShowMainDialog"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Detail Aplikasi"
                userToolTip="Detail Aplikasi"/>

      <itemData guid="379515da-830d-4150-9f70-17375552607c"
                icon="guid://379515da-830d-4150-9f70-17375552607c"
                dynamicCommand="MgoCorel.modSettings.Plugin_ShowSettingsForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Pengaturan"
                userToolTip="Pengaturan"/>

      <itemData guid="cb77d38a-f257-41ff-bebd-a5618aed5c19"
                icon="guid://cb77d38a-f257-41ff-bebd-a5618aed5c19"
                dynamicCommand="MgoCorel.modLabel.Plugin_ShowLabelForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Buat Label Spanduk"
                userToolTip="Buat Label Spanduk"/>

      <itemData guid="e3acc071-1633-4690-ab72-6d485e478b1f"
                icon="guid://e3acc071-1633-4690-ab72-6d485e478b1f"
                dynamicCommand="MgoCorel.modExport.Plugin_ExportLabel"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Export Spanduk"
                userToolTip="Export Spanduk"/>

      <itemData guid="e153e09d-7dbd-43fc-87dc-ee2f3f26c233"
                icon="guid://e153e09d-7dbd-43fc-87dc-ee2f3f26c233"
                dynamicCommand="MgoCorel.modSusun.Plugin_ShowSusunForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Susun Objek"
                userToolTip="Susun Objek"/>

      <itemData guid="73f734b6-a01a-4618-bbb0-e725d15837f5"
                icon="guid://73f734b6-a01a-4618-bbb0-e725d15837f5"
                dynamicCommand="MgoCorel.modImposisi.Plugin_ShowImposisiForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Imposisi Otomatis"
                userToolTip="Imposisi Otomatis"/>

      <itemData guid="4527162e-44c4-474c-b91b-3023d1072c5b"
                icon="guid://4527162e-44c4-474c-b91b-3023d1072c5b"
                dynamicCommand="MgoCorel.modRectangleNesting.Plugin_ShowRectangleNestingForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Rectangle Nesting"
                userToolTip="Rectangle Nesting"/>

      <itemData guid="8ad99ce3-0900-432f-a957-59a6b23e5a3e"
                icon="guid://8ad99ce3-0900-432f-a957-59a6b23e5a3e"
                dynamicCommand="MgoCorel.modNesting.Plugin_ShowNestingForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Nesting"
                userToolTip="Nesting"/>

      <itemData guid="8cdb5bfd-4cc6-421c-b5d8-79e447e0a253"
                icon="guid://8cdb5bfd-4cc6-421c-b5d8-79e447e0a253"
                dynamicCommand="MgoCorel.modHitung.Plugin_ShowHitungForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Hitung Harga"
                userToolTip="Hitung Harga"/>

      <itemData guid="11fcd242-68e2-4086-b1dd-0c94770774fc"
                icon="guid://11fcd242-68e2-4086-b1dd-0c94770774fc"
                dynamicCommand="MgoCorel.modDuplicate.Plugin_ShowDuplicateForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Duplicate Quantity"
                userToolTip="Duplicate Quantity"/>

      <itemData guid="d73dfff2-ed92-4037-8892-c240eeaf3da6"
                icon="guid://d73dfff2-ed92-4037-8892-c240eeaf3da6"
                dynamicCommand="MgoCorel.modNumbering.Plugin_ShowNumberingForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Numbering Otomatis"
                userToolTip="Numbering Otomatis"/>

      <itemData guid="ccdd8939-e735-436c-a223-2cc23200c453"
                icon="guid://ccdd8939-e735-436c-a223-2cc23200c453"
                dynamicCommand="MgoCorel.modJobBuilder.Plugin_ShowJobBuilderForm"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Job Builder"
                userToolTip="Job Builder"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/commandBars">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>

      <commandBarData guid="812bc65f-0465-470e-97b2-6139a95a0001"
                      type="menu"
                      captionRef="36f57221-bc31-4ddb-8c66-5c86b1f999b6"
                      userCreated="false"
                      userCaption="Detail"
                      nonLocalizableName="MgoCorel Detail"
                      flyout="true">
        <menu>
          <item dock="top" guidRef="38533386-1283-49fa-9e35-124f6fab8744"/>
          <item dock="top" guidRef="379515da-830d-4150-9f70-17375552607c"/>
        </menu>
      </commandBarData>

      <commandBarData guid="812bc65f-0465-470e-97b2-6139a95a0002"
                      type="menu"
                      captionRef="c64401b3-ea8a-4295-bbcb-d8c407dba836"
                      userCreated="false"
                      userCaption="Spanduk"
                      nonLocalizableName="MgoCorel Spanduk"
                      flyout="true">
        <menu>
          <item dock="top" guidRef="cb77d38a-f257-41ff-bebd-a5618aed5c19"/>
          <item dock="top" guidRef="e3acc071-1633-4690-ab72-6d485e478b1f"/>
        </menu>
      </commandBarData>

      <commandBarData guid="812bc65f-0465-470e-97b2-6139a95a0003"
                      type="menu"
                      captionRef="15f0179a-5594-4ec8-b3fd-69ee17eedcfa"
                      userCreated="false"
                      userCaption="Layout"
                      nonLocalizableName="MgoCorel Layout"
                      flyout="true">
        <menu>
          <item dock="top" guidRef="e153e09d-7dbd-43fc-87dc-ee2f3f26c233"/>
          <item dock="top" guidRef="73f734b6-a01a-4618-bbb0-e725d15837f5"/>
          <item dock="top" guidRef="4527162e-44c4-474c-b91b-3023d1072c5b"/>
          <item dock="top" guidRef="8ad99ce3-0900-432f-a957-59a6b23e5a3e"/>
        </menu>
      </commandBarData>

      <commandBarData guid="812bc65f-0465-470e-97b2-6139a95a0004"
                      type="menu"
                      captionRef="7f6d3bd7-7144-47a6-aaae-659ba4beb89d"
                      userCreated="false"
                      userCaption="Produksi"
                      nonLocalizableName="MgoCorel Produksi"
                      flyout="true">
        <menu>
          <item dock="top" guidRef="8cdb5bfd-4cc6-421c-b5d8-79e447e0a253"/>
          <item dock="top" guidRef="11fcd242-68e2-4086-b1dd-0c94770774fc"/>
          <item dock="top" guidRef="d73dfff2-ed92-4037-8892-c240eeaf3da6"/>
        </menu>
      </commandBarData>

      <commandBarData guid="812bc65f-0465-470e-97b2-6139a95a0005"
                      type="menu"
                      captionRef="4d9a55ff-0ad7-4268-8a7e-dd8fa602cc91"
                      userCreated="false"
                      userCaption="Job"
                      nonLocalizableName="MgoCorel Job"
                      flyout="true">
        <menu>
          <item dock="top" guidRef="ccdd8939-e735-436c-a223-2cc23200c453"/>
        </menu>
      </commandBarData>

      <xsl:if test="not(./commandBarData[@guid='a84d5e31-b1d6-4f8e-9c27-6274a0e11f35'])">
        <commandBarData guid="a84d5e31-b1d6-4f8e-9c27-6274a0e11f35"
                        userCreated="false"
                        userCaption="MgoCorel"
                        nonLocalizableName="MgoCorel"
                        type="toolbar">
          <toolbar type="toolbar" dock="top" itemFace="notSet" imageSize="medium">
            <item guidRef="36f57221-bc31-4ddb-8c66-5c86b1f999b6"/>
            <item guidRef="c64401b3-ea8a-4295-bbcb-d8c407dba836"/>
            <item guidRef="15f0179a-5594-4ec8-b3fd-69ee17eedcfa"/>
            <item guidRef="7f6d3bd7-7144-47a6-aaae-659ba4beb89d"/>
            <item guidRef="4d9a55ff-0ad7-4268-8a7e-dd8fa602cc91"/>
          </toolbar>
        </commandBarData>
      </xsl:if>

    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
