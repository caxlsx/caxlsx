# frozen_string_literal: true

module Axlsx
  # Theme represents the theme part of the package and is responsible for
  # generating the default Office theme that is required for encryption compatibility
  class Theme
    # The part name of this theme
    # @return [String]
    def pn
      THEME_PN
    end

    # Serializes the default theme to XML
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << <<~XML.delete("\n")
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
        <a:themeElements>
        <a:clrScheme name="Office">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
        <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
        <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
        <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
        <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
        <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
        <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
        <a:accent6><a:srgbClr val="F79646"/></a:accent6>
        <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
        <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>

        <a:fontScheme name="Office">
        <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        </a:majorFont>
        <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        </a:minorFont>
        </a:fontScheme>

        <a:fmtScheme name="Office">
        <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
        <a:gsLst>
        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
        <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
        <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        </a:gsLst>
        <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
        <a:gsLst>
        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
        <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        </a:gsLst>
        <a:lin ang="16200000" scaled="false"/>
        </a:gradFill>
        </a:fillStyleLst>

        <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill>
        <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:prstDash val="solid"/>
        </a:ln>
        </a:lnStyleLst>

        <a:effectStyleLst>
        <a:effectStyle>
        <a:effectLst>
        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="false">
        <a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr>
        </a:outerShdw>
        </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
        <a:effectLst>
        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="false">
        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
        </a:outerShdw>
        </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
        <a:effectLst>
        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="false">
        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
        </a:outerShdw>
        </a:effectLst>
        <a:scene3d>
        <a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera>
        <a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig>
        </a:scene3d>
        <a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>
        </a:effectStyle>
        </a:effectStyleLst>

        <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
        <a:gsLst>
        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
        </a:gsLst>
        <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
        <a:gsLst>
        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
        </a:gsLst>
        <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
        </a:bgFillStyleLst>
        </a:fmtScheme>
        </a:themeElements>

        <a:objectDefaults>
        <a:spDef>
        <a:spPr/>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:style>
        <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
        <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
        <a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
        </a:style>
        </a:spDef>
        <a:lnDef>
        <a:spPr/>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:style>
        <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
        <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
        <a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
        </a:style>
        </a:lnDef>
        </a:objectDefaults>
        <a:extraClrSchemeLst/>
        </a:theme>
      XML
    end
  end
end
