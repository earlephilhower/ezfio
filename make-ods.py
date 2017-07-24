#!/usr/bin/python

import argparse
import base64
import os
import sys
import tempfile
import zipfile


# Following routines emit the "standard" header and settings for an ODS sheet

def EmitXMLHeader():
    t  = '<?xml version="1.0" encoding="UTF-8"?>' + "\n"
    return t

def EmitOfficeDocumentContent():
    t  = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">'
    return t

def EmitOfficeScripts():
    t  = '<office:scripts/>'
    return t

def EmitOfficeFontFaceDecls():
    t  = '<office:font-face-decls>'
    t +=   '<style:font-face style:name="Liberation Sans1" svg:font-family="&apos;Liberation Sans1&apos;"/>'
    t +=   '<style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/>'
    t +=   '<style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-family-generic="system" style:font-pitch="variable"/>'
    t +=   '<style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/>'
    t += '</office:font-face-decls>'
    return t

def EmitStyleStyleColumn( name, width ):
    t  = '<style:style style:name="' + name + '" style:family="table-column">'
    t +=   '<style:table-column-properties fo:break-before="auto" style:column-width="' + width + '"/>'
    t += '</style:style>'
    return t

def EmitStyleStyleRow( name, height ):
    t  = '<style:style style:name="' + name + '" style:family="table-row">'
    t +=   '<style:table-row-properties style:row-height="' + height + '" fo:break-before="auto" style:use-optimal-row-height="true"/>'
    t += '</style:style>'
    return t

def EmitStyleStyleTable():
    t  = '<style:style style:name="ta1" style:family="table" style:master-page-name="mp1">'
    t +=   '<style:table-properties table:display="true" style:writing-mode="lr-tb"/>'
    t += '</style:style>'
    return t

def EmitNumberNumberStyle():
    t  = '<number:number-style style:name="N0">'
    t +=   '<number:number number:min-integer-digits="1"/>'
    t += '</number:number-style>'
    t += '<number:number-style style:name="N3">'
    t +=   '<number:number number:decimal-places="0" loext:min-decimal-places="0" number:min-integer-digits="1" number:grouping="true"/>'
    t += '</number:number-style>'
    t += '<number:percentage-style style:name="N10">'
    t +=   '<number:number number:decimal-places="0" loext:min-decimal-places="0" number:min-integer-digits="1"/>'
    t +=   '<number:text>%</number:text>'
    t += '</number:percentage-style>'
    return t

def EmitStyleStyleCe1Ce2():
    t  = '<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N0"/>'
    t +=  '<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N0">'
    t +=    '<style:table-cell-properties fo:background-color="#ffff00" style:text-align-source="fix" style:repeat-content="false" style:vertical-align="middle"/>'
    t +=    '<style:paragraph-properties fo:text-align="center"/>'
    t +=    '<style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>'
    t +=  '</style:style>'
    return t

def EmitStyleStyleCell( style, name ):
    t  = '<style:style style:name="' + style + '" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="' + name + '">'
    t +=   '<style:table-cell-properties fo:background-color="#ffff00"/>'
    t += '</style:style>'
    return t

def EmitStyleStyleGraphic():
    t  = '<style:style style:name="gr1" style:family="graphic">'
    t +=   '<style:graphic-properties draw:ole-draw-aspect="1"/>'
    t += '</style:style>'
    t += '<style:style style:name="gr2" style:family="graphic">'
    t +=   '<style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="left" draw:textarea-vertical-align="top" draw:auto-grow-height="false" draw:auto-grow-width="false" fo:padding-top="3.6pt" fo:padding-bottom="3.6pt" fo:padding-left="7.2pt" fo:padding-right="7.2pt" fo:wrap-option="no-wrap"/>'
    t += '</style:style>'
    return t

def EmitStyleStyleParagraph():
    t  = '<style:style style:name="P1" style:family="paragraph">'
    t +=   '<style:paragraph-properties fo:margin-left="0pt" fo:margin-right="0pt" fo:margin-top="0pt" fo:margin-bottom="0pt" fo:line-height="100%" fo:text-align="start" fo:text-indent="0pt" style:punctuation-wrap="hanging" style:writing-mode="lr-tb">'
    t +=     '<style:tab-stops/>'
    t +=   '</style:paragraph-properties>'
    t += '</style:style>'
    t += '<style:style style:name="P2" style:family="paragraph">'
    t +=   '<loext:graphic-properties draw:fill="none"/>'
    t +=   '<style:paragraph-properties style:writing-mode="lr-tb" style:font-independent-line-spacing="false"/>'
    t += '</style:style>'
    return t

def EmitStyleStyleText():
    t  = '<style:style style:name="T1" style:family="text">'
    t +=   '<style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:color="#000000" style:text-line-through-style="none" style:text-line-through-type="none" style:text-position="0% 100%" fo:font-family="Calibri" style:font-family-generic="roman" style:font-pitch="variable" fo:font-size="11pt" fo:letter-spacing="normal" fo:language="en" fo:country="US" fo:font-style="normal" style:text-underline-style="none" fo:font-weight="normal" style:text-underline-mode="continuous" style:text-overline-mode="continuous" style:text-line-through-mode="continuous" style:letter-kerning="false" style:font-family-asian="&apos;Segoe UI&apos;" style:font-pitch-asian="variable" style:font-size-asian="11pt" style:font-style-asian="normal" style:font-weight-asian="normal" style:font-family-complex="Tahoma" style:font-pitch-complex="variable" style:font-size-complex="11pt" style:font-style-complex="normal" style:font-weight-complex="normal"/>'
    t += '</style:style>'
    t += '<style:style style:name="T2" style:family="text">'
    t +=   '<style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:color="#000000" style:text-line-through-style="none" style:text-line-through-type="none" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="11pt" fo:letter-spacing="normal" fo:language="en" fo:country="US" fo:font-style="normal" style:text-underline-style="none" fo:font-weight="normal" style:text-underline-mode="continuous" style:text-overline-mode="continuous" style:text-line-through-mode="continuous" style:letter-kerning="true" style:font-family-asian="&apos;Segoe UI&apos;" style:font-pitch-asian="variable" style:font-size-asian="11pt" style:font-style-asian="normal" style:font-weight-asian="normal" style:font-family-complex="Tahoma" style:font-pitch-complex="variable" style:font-size-complex="11pt" style:font-style-complex="normal" style:font-weight-complex="normal"/>'
    t += '</style:style>'
    t += '<style:style style:name="T3" style:family="text">'
    t +=   '<style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:color="#000000" style:text-line-through-style="none" style:text-line-through-type="none" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="11pt" fo:letter-spacing="normal" fo:language="en" fo:country="US" fo:font-style="normal" style:text-underline-style="none" fo:font-weight="normal" style:text-underline-mode="continuous" style:text-overline-mode="continuous" style:text-line-through-mode="continuous" style:letter-kerning="true" style:font-family-asian="&apos;DejaVu Sans Mono&apos;" style:font-family-generic-asian="modern" style:font-pitch-asian="fixed" style:font-size-asian="11pt" style:font-style-asian="normal" style:font-weight-asian="normal" style:font-family-complex="&apos;DejaVu Sans Mono&apos;" style:font-family-generic-complex="modern" style:font-pitch-complex="fixed" style:font-size-complex="11pt" style:font-style-complex="normal" style:font-weight-complex="normal"/>'
    t += '</style:style>'
    return t

def EmitOfficeAutomaticStyles():
    t  = '<office:automatic-styles>'
    t += EmitStyleStyleColumn( "co1", "63.75pt" )
    t += EmitStyleStyleColumn( "co2", "54pt" )
    t += EmitStyleStyleColumn( "co3", "64.01pt" )
    t += EmitStyleStyleRow( "ro1", "14.26pt" )
    t += EmitStyleStyleRow( "ro2", "15pt" )
    t += EmitStyleStyleRow( "ro3", "13.8pt" )
    t += EmitStyleStyleRow( "ro4", "12.81pt" )
    t += EmitStyleStyleTable()
    t += EmitNumberNumberStyle()
    t += EmitStyleStyleCe1Ce2()
    t += EmitStyleStyleCell( "ce3", "N3" )
    t += EmitStyleStyleCell( "ce4", "N121" )
    t += EmitStyleStyleCell( "ce5", "N3" )
    t += EmitStyleStyleCell( "ce6", "N3" )
    t += EmitStyleStyleGraphic()
    t += EmitStyleStyleParagraph()
    t += EmitStyleStyleText()
    t += '</office:automatic-styles>'
    return t

def EmitOfficeBodySpreadsheetStart():
    t  = '<office:body>'
    t +=   '<office:spreadsheet>'
    t +=     '<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false"/>'
    return t


def EmitOfficeBodySpreadsheetEnd():
    t  =       '<table:named-expressions/>'
    t +=     '</office:spreadsheet>'
    t +=   '</office:body>'
    t += '</office:document-content>'
    return t

def EmitDrawFrameChart( zindex, name, width, height, x, y, objectid ):
    t  = '<draw:frame draw:z-index="' + str(zindex) + '" draw:name="' + str(name) + '" draw:style-name="gr1" svg:width="' + str(width) + '" svg:height="' + str(height) + '" svg:x="' + str(x) + '" svg:y="' + str(y) + '">'
    t +=   '<draw:object xlink:href="./Object ' + str(objectid) + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>'
#    t +=   '<draw:image xlink:href="./ObjectReplacements/Object ' + str(objectid) + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>'
    t += '</draw:frame>'
    return t


def EmitChartOfficeDocumentContent():
    t  = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:chartooo="http://openoffice.org/2010/chart" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">'
    return t

def EmitChartOfficeAutomaticStyles():
    t  = '<office:automatic-styles>'
    t += EmitNumberNumberStyle()
    t += '<style:style style:name="ch1" style:family="chart">'
    t +=   '<style:graphic-properties draw:stroke="none"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch2" style:family="chart">'
    t +=   '<style:chart-properties chart:auto-position="true" style:rotation-angle="0"/>'
    t +=   '<style:text-properties fo:color="#000000" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="13pt" style:font-size-asian="13pt" style:font-size-complex="13pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch3" style:family="chart">'
    t +=   '<style:chart-properties chart:auto-position="true"/>'
    t +=   '<style:graphic-properties svg:stroke-width="0.026cm" svg:stroke-color="#b3b3b3" draw:fill="none" draw:fill-color="#d9d9d9"/>'
    t +=   '<style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch4" style:family="chart">'
    t +=   '<style:chart-properties chart:symbol-type="automatic" chart:auto-position="true" chart:auto-size="true" chart:treat-empty-cells="leave-gap"/>'
    t +=   '<style:graphic-properties svg:stroke-width="0.026cm" svg:stroke-color="#b3b3b3" draw:fill="none"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch5" style:family="chart" style:data-style-name="N0">'
    t +=   '<style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false" chart:tick-marks-major-outer="false" chart:logarithmic="false" chart:reverse-direction="false" text:line-break="false" loext:try-staggering-first="false" chart:link-data-style-to-source="true" chart:axis-position="0" chart:axis-label-position="outside-start" chart:tick-mark-position="at-labels"/>'
    t +=   '<style:graphic-properties svg:stroke-width="0.018cm" svg:stroke-color="#b3b3b3"/>'
    t +=   '<style:text-properties fo:color="#000000" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch6" style:family="chart">'
    t +=   '<style:chart-properties chart:auto-position="true" style:rotation-angle="0"/>'
    t +=   '<style:text-properties fo:color="#000000" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="9pt" style:font-size-asian="9pt" style:font-size-complex="9pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch7" style:family="chart" style:data-style-name="N0">'
    t +=   '<style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false" chart:tick-marks-major-outer="false" chart:logarithmic="false" chart:origin="0" chart:gap-width="150" chart:reverse-direction="false" text:line-break="false" loext:try-staggering-first="false" chart:link-data-style-to-source="true" chart:axis-position="start"/>'
    t +=   '<style:graphic-properties svg:stroke-width="0.018cm" svg:stroke-color="#b3b3b3"/>'
    t +=   '<style:text-properties fo:color="#000000" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch8" style:family="chart">'
    t +=   '<style:chart-properties chart:auto-position="true" style:rotation-angle="90"/>'
    t +=   '<style:text-properties fo:color="#000000" style:text-position="0% 100%" fo:font-family="Calibri" fo:font-size="9pt" style:font-size-asian="9pt" style:font-size-complex="9pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch9" style:family="chart">'
    t +=   '<style:graphic-properties svg:stroke-width="0.018cm" svg:stroke-color="#b3b3b3"/>'
    t += '</style:style>'
    # Kludge alert! ch10 is the style we use for datasets, so set up a series of differently colored styles for use later
    for x in [ ['1', '000000'], ['2', 'ea2c2d'], ['3', '008c47'], ['4', '1859a9'], ['5', 'f37d22'], ['6', '662c91'], ['7', 'a11d20'], ['8', 'b33893'] ]:
        t += '<style:style style:name="ch10_' + x[0] + '" style:family="chart" style:data-style-name="N0">'
        t +=   '<style:chart-properties chart:symbol-type="named-symbol" chart:symbol-name="square" chart:symbol-width="0.25cm" chart:symbol-height="0.25cm" chart:link-data-style-to-source="true"/>'
        t +=   '<style:graphic-properties svg:stroke-width="0.08cm" svg:stroke-color="#' + x[1] + '" draw:fill="none" draw:fill-color="#' + x[1] + '"/>'
        t +=   '<style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>'
        t += '</style:style>'
    t += '<style:style style:name="ch11" style:family="chart" style:data-style-name="N0">'
    t += '<style:chart-properties chart:symbol-type="named-symbol" chart:symbol-name="square" chart:symbol-width="0.247cm" chart:symbol-height="0.247cm" chart:link-data-style-to-source="true"/>'
    t +=   '<style:graphic-properties svg:stroke-width="0.08cm" svg:stroke-color="#223344" draw:fill="none" draw:fill-color="#223344"/>'
    t +=   '<style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch12" style:family="chart">'
    t +=   '<style:graphic-properties draw:stroke="solid" svg:stroke-width="0.026cm" svg:stroke-color="#b3b3b3" draw:fill="none" draw:fill-color="#d9d9d9"/>'
    t += '</style:style>'
    t += '<style:style style:name="ch13" style:family="chart">'
    t +=   '<style:graphic-properties draw:fill-color="#d9d9d9"/>'
    t += '</style:style>'
    t += '</office:automatic-styles>'
    return t


def EmitLineChartBody( width, height, titlex, titley, titletext, legend, legendx, legendy, plotx, ploty, plotw, ploth, xaxisx, xaxisy, xaxistext, yaxisx, yaxisy, yaxistext, xcatrange, dataserieslist ):
    t  = '<office:body>'
    t +=   '<office:chart>'
    t +=   '<chart:chart svg:width="' + width + '" svg:height="' + height + '" xlink:href=".." xlink:type="simple" chart:class="chart:line" chart:style-name="ch1">'
    t +=   '<chart:title svg:x="' + titlex +'" svg:y="' + titley + '" chart:style-name="ch2">'
    t +=     '<text:p>' + titletext + '</text:p>'
    t +=   '</chart:title>'
    if legend:
        t += '<chart:legend chart:legend-position="bottom" svg:x="' + legendx + '" svg:y="' + legendy + '" style:legend-expansion="wide" chart:style-name="ch3"/>'
    # Removed table:cell-range-address, hope this is recalculated properly!
    t +=   '<chart:plot-area chart:style-name="ch4" chart:data-source-has-labels="column" svg:x="' + plotx + '" svg:y="' + ploty + '" svg:width="' + plotw +'" svg:height="' + ploth + '">'
#					<chartooo:coordinate-region svg:x="2.358cm" svg:y="1.527cm" svg:width="13.042cm" svg:height="4.884cm"/>
    t += '<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5" chartooo:axis-type="text">'
    t +=   '<chart:title svg:x="' + xaxisx + '" svg:y="' + yaxisy + '" chart:style-name="ch6">'
    t +=     '<text:p>' + xaxistext + '</text:p>'
    t +=   '</chart:title>'
    t +=   '<chart:categories table:cell-range-address="' + xcatrange + '"/>'
    t += '</chart:axis>'
    t += '<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch7">'
    t +=   '<chart:title svg:x="' + yaxisx + '" svg:y="' + yaxisy + '" chart:style-name="ch8">'
    t +=     '<text:p>' + yaxistext + '</text:p>'
    t +=   '</chart:title>'
    t +=   '<chart:grid chart:style-name="ch9" chart:class="major"/>'
    t += '</chart:axis>'
    sernum = 1
    for series in dataserieslist:
        t += '<chart:series loext:label-string="' + str(series['label']) + '" chart:style-name="ch10_' + str(sernum) + '" chart:values-cell-range-address="' + str(series['data']) + '" chart:class="chart:line">'
#        t +=   '<chart:data-point chart:repeated="9"/>'
        t += '</chart:series>'
        sernum = sernum + 1
    t += '<chart:wall chart:style-name="ch12"/>'
    t += '<chart:floor chart:style-name="ch13"/>'
    t += '</chart:plot-area>'
    t += '</chart:chart>'
    t += '</office:chart>'
    t += '</office:body>'
    t += '</office:document-content>'
    return t

def Meta():
    t  = EmitXMLHeader()
    t += '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:chartooo="http://openoffice.org/2010/chart" office:version="1.2">'
    t += '<office:meta>'
    t += '<meta:generator>ezFIO</meta:generator>'
    t += '</office:meta>'
    t += '</office:document-meta>'
    return t

def Styles():
    t  = EmitXMLHeader()
    t += '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:chartooo="http://openoffice.org/2010/chart" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">'
    t +=   '<office:styles/>'
    t += '</office:document-styles>'
    return t

def Mimetype():
    t = 'application/vnd.oasis.opendocument.spreadsheet'
    return t

def Chart( width, height, titlex, titley, titletext, legend, legendx, legendy, plotx, ploty, plotw, ploth, xaxisx, xaxisy, xaxistext, yaxisx, yaxisy, yaxistext, xcatrange, dataserieslist ):
    t  = EmitXMLHeader()
    t += EmitChartOfficeDocumentContent()
    t += EmitChartOfficeAutomaticStyles()
    t += EmitLineChartBody( width, height, titlex, titley, titletext, legend, legendx, legendy, plotx, ploty, plotw, ploth, xaxisx, xaxisy, xaxistext, yaxisx, yaxisy, yaxistext, xcatrange, dataserieslist )
    return t

def SpreadsheetStart():
    t  = EmitXMLHeader()
    t += EmitOfficeDocumentContent()
    t += EmitOfficeScripts()
    t += EmitOfficeFontFaceDecls()
    t += EmitOfficeAutomaticStyles()
    t += EmitOfficeBodySpreadsheetStart()
    return t

def SpreadsheetEnd():
    t  = EmitOfficeBodySpreadsheetEnd()
    return t

def ManifestRDF():
    t  = EmitXMLHeader()
    t += '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
    t +=   '<rdf:Description rdf:about="styles.xml">'
    t +=     '<rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#StylesFile"/>'
    t +=   '</rdf:Description>'
    t +=   '<rdf:Description rdf:about="">'
    t +=     '<ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="styles.xml"/>'
    t +=   '</rdf:Description>'
    t +=   '<rdf:Description rdf:about="content.xml">'
    t +=     '<rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>'
    t +=   '</rdf:Description>'
    t +=   '<rdf:Description rdf:about="">'
    t +=     '<ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="content.xml"/>'
    t +=   '</rdf:Description>'
    t +=   '<rdf:Description rdf:about="">'
    t +=      '<rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>'
    t +=   '</rdf:Description>'
    t += '</rdf:RDF>'
    return t


def Manifest( numCharts ):
    t  = EmitXMLHeader()
    t += '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
    t +=   '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
    for i in range(1, numCharts + 1):
        t += '<manifest:file-entry manifest:full-path="Object ' + str(i) + '/meta.xml" manifest:media-type="text/xml"/>'
        t += '<manifest:file-entry manifest:full-path="Object ' + str(i) + '/styles.xml" manifest:media-type="text/xml"/>'
        t += '<manifest:file-entry manifest:full-path="Object ' + str(i) + '/content.xml" manifest:media-type="text/xml"/>'
        t += '<manifest:file-entry manifest:full-path="Object ' + str(i) + '/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>'
    t +=   '<manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>'
    t +=   '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
    t +=   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
    t += '</manifest:manifest>'
    return t


def EmitSheetFromCSV(sheetName, csvName):
    t  = '<table:table table:name='
    t += '"' + sheetName + '"' + ' table:style-name="ta1" > '
    t += '<table:table-column table:style-name="co1" '
    t += 'table:default-cell-style-name="Default"/>'
    # Insert the rows, one entry at a time
    with open(csvName) as f:
        for line in f:
            line = line.rstrip()
            t += '<table:table-row table:style-name="ro1">'
            for val in line.split(','):
                try:
                    cell  = '<table:table-cell office:value-type="float" '
                    cell += 'office:value="' + str(float(val))
                    cell += '" calcext:value-type="float"><text:p>'
                    cell += str(float(val)) + '</text:p></table:table-cell>'
                except: # It's not a float, so let's call it a string
                    cell  = '<table:table-cell office:value-type="string" '
                    cell += 'calcext:value-type="string"><text:p>'
                    cell += str(val) + '</text:p></table:table-cell>'
                t += cell
            t += '</table:table-row>'
        f.close()
    # Close the tags
    t += '</table:table>'
    return t

def EmitHeaderRows(inputFiles):
    for item in [ 'Label', 'Drive', 'Model', 'Serial', 'AvailCap', 'TestCap', 'CPU', 'Cores', 'Freq', 'OS', 'FIO' ]:
        t += '<table:table-row table:style-name="ro1">'
        t +=   '<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>' + item + '</text:p></table:table-cell>'
        for elem in inputFiles:
            t += '<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>' + elem[item] + '</text:p></table:table-cell>'
        t += '</table:table-row>'
    return t

def EmitGraphSheet(inputFiles):
    t  = '<table:table table:name='
    t += '"' + "Graphs" + '"' + ' table:style-name="ta1" > '
    t += '<table:shapes>'
    t += EmitDrawFrameChart( "1", "first chart", "11cm", "10cm", "1cm", "2cm", 1 )
    t += '</table:shapes>'
    t += '<table:table-column table:style-name="co1" table:number-columns-repeated="100" table:default-cell-style-name="ce1"/>'
    t += EmitHeaderRows(inputFiles)
    t += '</table:table>'
    return t

parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter,
    description="Combine multiple EZFIO runs into a single, labeled ODS file.")
parser.add_argument("--dir", "-d", dest = "dir", default=[], help="Directory to add to sheet (multiple options allowed)", required=True)
args = parser.parse_args()




odsfile = "out.ods"
if os.path.exists(odsfile):
    os.unlink(odsfile)

mimetypezip = """
UEsDBAoAAAAAAOKbNUiFbDmKLgAAAC4AAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub2Fz
aXMub3BlbmRvY3VtZW50LnNwcmVhZHNoZWV0UEsBAj8ACgAAAAAA4ps1SIVsOYouAAAALgAAAAgA
JAAAAAAAAACAAAAAAAAAAG1pbWV0eXBlCgAgAAAAAAABABgAAAyCUsVU0QFH/eNMmlTRAUf940ya
VNEBUEsFBgAAAAABAAEAWgAAAFQAAAAAAA==
"""
with open(odsfile, 'wb') as f:
    f.write(base64.b64decode( mimetypezip ))

ods = zipfile.ZipFile(odsfile, "a", zipfile.ZIP_DEFLATED)
ods.writestr( "manifest.rdf", ManifestRDF() )
ods.writestr( "meta.xml", Meta() )
ods.writestr( "META-INF/manifest.xml", Manifest(1) )
ods.writestr( "content.xml", SpreadsheetStart() + EmitGraphSheet() + EmitSheetFromCSV("data", "a.csv") + SpreadsheetEnd() )
ods.writestr( "Object 1/meta.xml", Meta() )
ods.writestr( "Object 1/styles.xml", Styles() )
ods.writestr( "Object 1/content.xml", Chart( "11cm", "10cm", "10pt", "20pt", "Cats and Dogs", True, "30pt", "40pt", "100pt", "100pt", "90pt", "80pt", "4pt", "5pt", "xaxishere", "6pt", "7pt", "yaxistext", "data.A2:data.A4", [ {'label' : "cat1", 'data' : "data.B2:data.B4"}, {'label': "cat2", 'data' : "data.C2:data.C4" } ] ) )

ods.close()
