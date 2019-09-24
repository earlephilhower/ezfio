#!/usr/bin/python

# ezfio 1.0
# earle.philhower.iii@hgst.com
#
# ------------------------------------------------------------------------
# ezfio is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# ezfio is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with ezfio.  If not, see <http://www.gnu.org/licenses/>.
# ------------------------------------------------------------------------
#
# Usage:   ./append.py --source <old.ods> --append <new.ods> --suffix <_new> --color <223344> --output <combined.ods>


import argparse
import base64
import datetime
import json
import os
import platform
import pwd
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import zipfile

def ParseArgs():
    """Parse command line options into globals."""
    global sourceODS, appendODS, destODS, suffix, color
    parser = argparse.ArgumentParser(
                 formatter_class=argparse.RawDescriptionHelpFormatter,
    description="A tool to add a dataset to an existing ezFIO ODS file.",
    epilog="")
    parser.add_argument("--source", "-s", dest = "sourceODS",
        help="First ODS file with 1 or more test runs included", required=True)
    parser.add_argument("--append", "-a", dest="appendODS",
        help="ODS file with tests to append to source", required=True)
    parser.add_argument("--suffix", "-x", dest="suffix",
        help="Suffix to append to data tables from appended ODS", required=True)
    parser.add_argument("--color", "-c", dest="color",
        help="Color to use for graphed data in appended ODS (rrggbb format)", required=True)
    parser.add_argument("--output", "-o", dest="destODS",
        help="Location where results should be saved", required=True)
    args = parser.parse_args()
    sourceODS = args.sourceODS
    appendODS = args.appendODS
    destODS = args.destODS
    suffix = args.suffix
    color = args.color

def GenerateCombinedODS():
    """Builds a new ODS spreadsheet w/graphs from generated test CSV files."""

    def GetContentXMLFromODS( odssrc ):
        """Extract content.xml from an ODS file, where the sheet lives."""
        ziparchive = zipfile.ZipFile( odssrc )
        content = ziparchive.read("content.xml")
        content = content.replace("\n", "")
        return content

    def CSVtoXMLSheet(sheetName, csvName):
        """Replace a named sheet with the contents of a CSV file."""
        newt  = '<table:table table:name='
        newt += '"' + sheetName + '"' + ' table:style-name="ta1" > '
        newt += '<table:table-column table:style-name="co1" '
        newt += 'table:default-cell-style-name="Default"/>'
        # Insert the rows, one entry at a time
        with open(csvName) as f:
            for line in f:
                line = line.rstrip()
                newt += '<table:table-row table:style-name="ro1">'
                for val in line.split(','):
                    try:
                        cell  = '<table:table-cell office:value-type="float" '
                        cell += 'office:value="' + str(float(val))
                        cell += '"><text:p>'
                        cell += str(float(val)) + '</text:p></table:table-cell>'
                    except: # It's not a float, so let's call it a string
                        cell  = '<table:table-cell office:value-type="string" '
                        cell += '><text:p>'
                        cell += str(val) + '</text:p></table:table-cell>'
                    newt += cell
                newt += '</table:table-row>'
            f.close()
        # Close the tags
        newt += '</table:table>'
        return newt

    def AppendSheetFromCSV(sheetName, csvName, xmltext):
        """Add a new sheet to the XML from the CSV file."""
        newt = CSVtoXMLSheet(sheetName, csvName)

        # Replace the XML using lazy string matching
        searchstr  = '<table:named-expressions/>'
        return re.sub(searchstr, newt + searchstr, xmltext)

    def UpdateContentXMLToODS_text( odssrc, odsdest, xmltext ):
        """Replace content.xml in an ODS w/an in-memory copy and write new.

        Replace content.xml in an ODS file with in-memory, modified copy and
        write new ODS. Can't just copy source.zip and replace one file, the
        output ZIP file is not correct in many cases (opens in Excel but fails
        ODF validation and LibreOffice fails to load under Windows).

        Also strips out any binary versions of objects and the thumbnail,
        since they are no longer valid once we've changed the data in the
        sheet.
        """
        global suffix

        if os.path.exists(odsdest):
            os.unlink(odsdest)

        # Windows ZipArchive will not use "Store" even with "no compression"
        # so we need to have a mimetype.zip file encoded below to match spec:
        mimetypezip = """
UEsDBAoAAAAAAOKbNUiFbDmKLgAAAC4AAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub2Fz
aXMub3BlbmRvY3VtZW50LnNwcmVhZHNoZWV0UEsBAj8ACgAAAAAA4ps1SIVsOYouAAAALgAAAAgA
JAAAAAAAAACAAAAAAAAAAG1pbWV0eXBlCgAgAAAAAAABABgAAAyCUsVU0QFH/eNMmlTRAUf940ya
VNEBUEsFBgAAAAABAAEAWgAAAFQAAAAAAA==
"""
        zipbytes = base64.b64decode( mimetypezip )
        with open(odsdest, 'wb') as f:
            f.write(zipbytes)

        zasrc = zipfile.ZipFile(odssrc, 'r')
        zadst = zipfile.ZipFile(odsdest, 'a', zipfile.ZIP_DEFLATED)
        for entry in zasrc.namelist():
            if entry == "mimetype":
                continue
            elif entry.endswith('/') or entry.endswith('\\'):
                continue
            elif entry == "content.xml":
                zadst.writestr( "content.xml", xmltext)
            elif ("Object" in entry) and ("content.xml" in entry):
                # Remove <table:table table:name="local-table"> table
                rdbytes = zasrc.read(entry)
                outbytes = re.sub('<table:table table:name="local-table">.*</table:table>', "", rdbytes)
                # Add in extra chart series following existing format...
                searchStr = '<chart:series .*</chart:series>'
                match = re.search(searchStr, outbytes);
                addl = ""
                if match:
                    fmt = match.group(0)
                    addl = fmt;
                    for sheet in [ "Tests", "Timeseries", "Exceedance"]:
                        addl = re.sub( sheet, sheet+suffix, addl )
                    # Remove any existing label and add updated one
                    addl = re.sub("loext:label-string=\".*?\"" , "", addl );
                    addl = re.sub ("<chart:series ", "<chart:series " + "loext:label-string=\""+suffix+"\" ", addl) 
                    styleMatch = re.search("chart:style-name=\"(.)*?\"", fmt)
                    if styleMatch:
                        styleName = re.sub("chart:style-name=\"", "", styleMatch.group(0) )
                        styleName = re.sub("\".*", "", styleName)
                        # Change the style requested in new one...
                        addl = re.sub( "\"" + styleName + "\"", "\"" + styleName + suffix + "\"", addl )
                        # And patch in the new chart:series entry
                        outbytes = re.sub ( searchStr, fmt + addl, outbytes )
                        # Now make the new style...
                        oldStyleMatch = re.search( "<style:style style:name=\"" + styleName + ".*?</style:style>" , outbytes )
                        if oldStyleMatch:
                            oldStyle = oldStyleMatch.group(0)
                            newStyle = re.sub( "\"" + styleName + "\"", "\"" + styleName + suffix + "\"", oldStyle)
                            # Change the embedded color:
                            newStyle = re.sub( "svg:stroke-color=\"#.*?\"", "svg:stroke-color=\"#" + color + "\"", newStyle )
                            # Add in the new style...
                            outbytes = re.sub ( oldStyle, oldStyle + newStyle, outbytes )
                        # Add legend if it doesn't exist
                        legendMatch = re.search("<chart:legend .*?/>", outbytes)
                        if not legendMatch:
                            # Put in hardcoded one...looks like junk, but can be tweaked by user in application
                            legend = "<chart:legend chart:legend-position=\"bottom\" svg:x=\"0.000cm\" svg:y=\"0.000cm\" style:legend-expansion=\"wide\" chart:style-name=\"ch3\"/>";
                            outbytes = re.sub ("</chart:title>", "</chart:title>" + legend, outbytes )
                zadst.writestr(entry, outbytes)
            elif entry == "META-INF/manifest.xml":
                # Remove ObjectReplacements from the list
                rdbytes = zasrc.read(entry)
                outbytes = ""
                lines = rdbytes.split("\n")
                for line in lines:
                    if not ( ("ObjectReplacement" in line) or ("Thumbnails" in line) ):
                        outbytes = outbytes + line + "\n"
                zadst.writestr(entry, outbytes)
            elif ("Thumbnails" in entry) or ("ObjectReplacement" in entry):
                # Skip binary versions
                continue
            else:
                rdbytes = zasrc.read(entry)
                zadst.writestr(entry, rdbytes)
        zasrc.close()
        zadst.close()


    global sourceODS, appendODS, destODS
    
    # First rename and append the extra data sheets
    xmlsrc = GetContentXMLFromODS( sourceODS )
    xmlapp = GetContentXMLFromODS( appendODS )
    for tableName in [ "Tests", "Timeseries", "Exceedance" ]:
        searchStr = '<table:table table:name="' + tableName + '".*?</table:table>'
        sheetMatch = re.search(searchStr, xmlapp);
        if sheetMatch:
            sheet = sheetMatch.group(0)
            # Rename the table
            sheet = re.sub( '"' + tableName + '"', '"' + tableName + suffix + '"', sheet);
            # Stick it right before the end of the list
            searchStr  = '<table:named-expressions/>'
            xmlsrc = re.sub(searchStr, sheet + searchStr, xmlsrc)
    UpdateContentXMLToODS_text( sourceODS, destODS, xmlsrc )

sourceODS = ""
appendODS = ""
destODS = ""
suffix = ""
color = ""

if __name__ == "__main__":
    ParseArgs()
    GenerateCombinedODS()

