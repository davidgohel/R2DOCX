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
# simple paragraphs
doc <- addParagraph( doc, value = c("Hello!", "How are you today?"), stylename = "Normal")


# define 2 paragraph of text to add later
x = c( "[animal] eat [food].", "tigers eat [animal]." )
# define styles that will be used for formating replacements texts
repl.styles = list(
		animal= textProperties( font.size = 12, font.family="Courier New", color="red" )
		, food= textProperties( font.size = 12, font.family="Courier New", color="blue" )
)
# define replacements texts
repl = list( animal = "buffalos" , food = "grass" )

doc <- addParagraph( doc, value = x, stylename = "Normal"
		, replacements = repl
		, replacement.styles = repl.styles
)

writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
