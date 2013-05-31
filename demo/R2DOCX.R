# TODO: Documentation de R2DOCX.R
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 24 mai 2013
# Version: 0.1
###############################################################################
require( R2DOCX )

# Word document to write
docx.file <- "document.docx"

# Remove file if it already exists
if(file.exists( docx.file )) 
	file.remove( docx.file )


header.cellProperties = cellProperties( border.left.width = 0, border.right.width = 0
		, border.bottom.width = 3, border.top.width = 0
		, padding = 2 )
data.cellProperties = cellProperties( border.left.width = 1, border.right.width = 1
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






doc <- new("Docx", title = "My example" )
styles(doc)
doc <- setHeaderStyle(doc
		, stylenames = c("Titre1", "Titre2", "Titre3"
				, "Titre4", "Titre5", "Titre6"
				, "Titre7", "Titre8", "Titre9" ) )
doc <- addHeader(doc, "Table of content", 1);
doc <- addTOC(doc)
doc <- addHeader(doc, "Iris dataset", 1);
doc = addTable( doc, iris, formats = get.default.formats() )

doc <- addHeader(doc, "Text examples", 1);
doc <- addParagraph( doc, value = "Normal text", par.style = "Normal" )
doc <- addParagraph( doc, value = "Citation text", par.style = "Citation" )
		
repl.styles = list(
		animal= textProperties(font.size = 12, font.family="Courier New", color="red" )
		, food= textProperties(font.size = 12, font.family="Courier New", color="blue" )
)

doc <- addParagraph( doc, value = "[animal] eats [food].", par.style = "Normal"
		, replacements = list( animal = "donkey" , food = "grass" )
		, replacement.styles = repl.styles
)
doc <- addHeader(doc, "Table example with formating", 1);
require(lattice)
data = barley[, c("variety", "site", "year", "yield" )]
data$PC = runif( nrow( data ), 0, 1 )# create dummy percent


doc <- addTable( doc
		, data
		, formats = my.formats
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
doc <- addHeader(doc, "Graph example", 1);
require( ggplot2 )
doc = addPlot( doc
		, function(){
			print( qplot(Sepal.Length, Petal.Length, data = iris, color = Species, size = Petal.Width, alpha = I(0.7)) )
		}
		, width = 10, height = 8, legend="graph example" 
		, stylename="PlotReference", pointsize=10
)

writeDoc( doc, docx.file )


if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}

