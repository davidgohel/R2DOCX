# TODO: Documentation de R2DOCX.R
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 24 mai 2013
# Version: 0.1
###############################################################################
require( R2DOCX )
data( weights.summary )
# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )


header.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 1, border.top.width = 0
		, padding = 3 )
data.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 1, border.top.width = 0
		, padding = 2 )

header.textProperties = textProperties( font.size = 10, font.weight = "bold" )
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
		, data.cell = data.cellProperties
		, data.text = data.textProperties
)






doc <- new("Docx", title = "My example" )
styles(doc)
doc <- addParagraph( doc, value = "My example", stylename = "Titre" )
doc = addPageBreak( doc )

doc <- addHeader(doc, "Table of content", 1);
doc <- addTOC(doc)
doc = addPageBreak( doc )
doc <- addHeader(doc, "Iris dataset", 1);
doc = addTable( doc, head( iris )	, formats = my.formats )

doc = addPageBreak( doc )
doc <- addHeader(doc, "Text examples", 1);
doc <- addParagraph( doc, value = "Normal text", stylename = "Normal" )
doc <- addParagraph( doc, value = "Citation text", stylename = "Citation" )
		
repl.styles = list(
		animal= textProperties(font.size = 12, font.family="Courier New", color="red" )
		, food= textProperties(font.size = 12, font.family="Courier New", color="blue" )
)

doc <- addParagraph( doc, value = c( "[animal] eat [food].", "tigers eat [animal]." )
	, stylename = "Normal"
	, replacements = list( animal = "donkey" , food = "grass" )
	, replacement.styles = repl.styles
)
doc = addPageBreak( doc )
doc <- addHeader(doc, "Table example with formating", 1);

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

doc = addPageBreak( doc )
doc <- addHeader(doc, "Graph example", 1);
doc = addPlot( doc
	, fun = plot
	, x = rnorm( 100 )
	, y = rnorm (100 )
	, main = "base plot main title"
)
writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}

