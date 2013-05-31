#############################################################################
#
# R2DOCX
# Copyright (c) 2013, David Gohel All rights reserved.
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################

# Author: David GOHEL <david.gohel@gmail.com>
# Date: 19 mai 2013
# Version: 0.1
###############################################################################

require( R2DOCX )

# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )

# create document
doc <- new("Docx", title = "My example" )
styles(doc)
#[1] "Normal"                  "Titre1"                  "Titre2"                 
#[4] "Titre3"                  "Titre4"                  "Titre5"                 
#[7] "Titre6"                  "Titre7"                  "Titre8"                 
#[10] "Titre9"                  "Policepardfaut"          "TableauNormal"  
#...
#[31] "Citation"                "CitationCar"             "Citationintense"
doc <- setHeaderStyle(doc
		, stylenames = c("Titre1", "Titre2", "Titre3"
				, "Titre4", "Titre5", "Titre6"
				, "Titre7", "Titre8", "Titre9" ) )

doc <- addTOC(doc)


doc <- addHeader(doc, "Title 1", 1);
doc <- addHeader(doc, "Title 1.1", 2);
doc <- addHeader(doc, "Title 1.1.1", 3);
doc <- addHeader(doc, "Title 1.1.1.1", 4);

doc <- addParagraph( doc, value = "Normal text", par.style = "Normal" )
doc <- addParagraph( doc, value = "Citation text", par.style = "Citation" )
writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
