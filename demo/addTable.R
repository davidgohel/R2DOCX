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
require(lattice)

# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )

# create document
doc <- new("Docx", title = "My example" )

doc <- addParagraph( doc, value = "iris dataset :", par.style = "Normal" )

doc <- addTable( doc, iris )

doc <- addParagraph( doc, value = "another table :", par.style = "Normal" )
data = barley[, c("variety", "site", "year", "yield" )]
data$PC = runif( nrow( data ), 0, 1 )# create dummy percent


doc <- addTable( doc
		, data
		, header.labels = list("variety" = "Variety", "site" = "Site"
				, "year" = "Year", "yield" = "Yield", "PC" = "Percent column")
		, grouped.cols=list( "Types" = c( "variety", "site" )
				, "year" = "year"
				, "Mesures" = c("yield", "PC") )
		, span.first.columns = 1
		, col.types = list("variety" = "character", "site" = "character"
				, "year" = "integer", "yield" = "double", "PC"="percent" ) 
		, col.fontcolors = list( "PC" = ifelse( data$PC < 0.05 , "#FF0000", "#000000") )
)


writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
