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


header.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 2, border.top.width = 0
		, padding = 2 )
data.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 1, border.top.width = 0
		, padding = 1 )

header.textProperties = textProperties( font.size = 12, color = "gray"
		, font.weight = "bold" )
data.textProperties = textProperties( font.size = 10 )

leftPar = parProperties( text.align = "left", padding.left = 4, padding.right = 4)
rightPar = parProperties( text.align = "right", padding.left = 4, padding.right = 4)
centerPar = parProperties( text.align = "center", padding.left = 4, padding.right = 4)

my.formats = tableProperties( 
		character.par = leftPar
		, percent.par = rightPar
		, double.par = rightPar
		, integer.par = rightPar
		, groupedheader.par = centerPar
		, groupedheader.text = header.textProperties
		, groupedheader.cell = header.cellProperties
		, header.cell = header.cellProperties
		, header.par = centerPar
		, header.text = header.textProperties
		, double.cell = data.cellProperties
		, integer.cell = data.cellProperties
		, percent.cell = data.cellProperties
		, character.cell = data.cellProperties
		, character.text = data.textProperties
		, double.text = data.textProperties
		, percent.text = data.textProperties
		, integer.text = data.textProperties
)

doc <- addTable( doc, iris, formats = my.formats )

writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
