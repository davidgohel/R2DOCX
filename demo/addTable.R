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

data( measured.weights )
data( weights.summary )

# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )

# create document
doc <- new("Docx", title = "My example" )

# add a title
doc <- addHeader( doc, "Table example", 1 )

# add the first 5 lines of measured.weights in the docx
doc <- addTable( doc
	, head( measured.weights, n = 5 )
	, span.columns = c( "id" ) 
	)

# add another title
doc <- addHeader( doc, "Another table example", 1 )

# add and format weights.summary in the docx
doc <- addTable( doc
		, data = weights.summary
		, header.labels = list("id" = "Subject Identifier", "avg.weight" = "Average Weight"
				, "regularity" = "Regularity", "visit.n" = "Number of visits", "last.day" = "Last visit")
		, grouped.cols=list( "id" = "id"
				, "Summary" = c( "avg.weight",  "regularity", "visit.n", "last.day" ) )
		, col.types = list( "id" = "character", "avg.weight" = "double"
				, "regularity" = "percent", "visit.n" = "integer"
				, "last.day" = "date" ) 
		, col.fontcolors = list( "regularity" = ifelse( weights.summary$regularity < 0.5 , "gray", "blue") )
)

# write the docx onto the disk
writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
