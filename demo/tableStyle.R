# Author: David GOHEL <david.gohel@gmail.com>
# Date: 19 mai 2013
# Version: 0.1
###############################################################################

require( R2DOCX )
data( weights.summary )
# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )

# create document
doc <- new("Docx", title = "My example" )


header.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 2, border.top.width = 0
		, padding = 5, background.color="#30315f" )
data.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 1, border.top.width = 0
		, padding = 3, background.color="#eeeeee"  )

header.textProperties = textProperties( font.size = 12, color = "#c13100", font.weight = "bold" )
data.textProperties = textProperties( font.size = 10, color = "#303030" )

leftPar = parProperties( text.align = "left" )
rightPar = parProperties( text.align = "right" )
centerPar = parProperties( text.align = "center" )

my.formats = tableProperties( 
		character.par = leftPar, percent.par = rightPar
		, double.par = rightPar, integer.par = rightPar
		, groupedheader.par = centerPar, groupedheader.text = header.textProperties
		, groupedheader.cell = header.cellProperties
		, header.cell = header.cellProperties, header.par = centerPar
		, header.text = header.textProperties
		, data.cell = data.cellProperties, data.text = data.textProperties
		, percent.addsymbol = " %"
		, integer.digit = 1L
		, fraction.double.digit = 3L
		, fraction.percent.digit = 2L
)




doc <- new("Docx", title = "My example" )
doc <- addTable( doc, weights.summary
		, formats = my.formats
		, col.types = list( "id" = "character", "avg.weight" = "double"
				, "regularity" = "percent", "visit.n" = "integer"
				, "last.day" = "date" ) 
)
doc <- addLineBreak( doc )
doc <- addTable( doc, weights.summary
		, formats = tableProperties( header.cell = cellProperties( background.color = "gray") )
		, col.types = list( "id" = "character", "avg.weight" = "double"
				, "regularity" = "percent", "visit.n" = "integer"
				, "last.day" = "date" ) 
)
writeDoc( doc, docx.file )

if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
