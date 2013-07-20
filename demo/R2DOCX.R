# TODO: Documentation de R2DOCX.R
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 07 june 2013
# Version: 0.1
###############################################################################
require( R2DOCX )
require(ggplot2)

# The 2 datasets that will be used
data( weights.summary )
data( measured.weights )

# Word document to write
docx.file = "document.docx"
# Word document to use as base document or a template
template.file = file.path( find.package("R2DOCX"), "templates/TEMPLATE_01.docx", fsep = "/" )

# Create a new object
doc = new("Docx"
		, title = "My example" 
		, basefile = template.file
)
# in the base doc, 2 bookmarks are existing - let's replace them by some values
doc = addParagraph( doc, c( "John Doe", "Tarzan" ), stylename = "Normal", bookmark="AUTHOR" )
doc = addParagraph( doc, date(), stylename = "Normal", bookmark="DATE"  )

# in the base doc, style 'TitleDoc' is existing
# we use it to generate a title ("My example") for the doc
doc = addParagraph( doc, value = "My example", stylename = "TitleDoc" )
# add a page break
doc = addPageBreak( doc )


# add a page for TOC (user will be asked to update this TOC 
# when document will be opened the first time)
doc = addHeader(doc, "Table of content", 1);
doc = addTOC(doc)
doc = addPageBreak( doc )



# add a texts (normal and bullets)
doc = addHeader(doc, "Dataset presentation", 1);
doc = addParagraph( doc, value = "Dataset 'measured.weights' is set of weights measurements collected on few people before savate training.", stylename = "Normal" )
texts = c( "Column 'gender' is the subject gender." 
		, "Column 'id' is the unique subject identifier." 
		, "Column 'cat' is the subject weight category." 
		, "Column 'competitor' tells wether or not subject practice competition." 
		, "Column 'day' is the day of measurement." 
		, "Column 'weight' is the measured weight." 
)
doc = addParagraph( doc, value = texts, stylename = "BulletList" )


# simple addTable
doc = addHeader(doc, "Dataset", 1);
doc = addHeader(doc, "Dataset first lines", 2);
doc = addTable( doc
		, data = head( measured.weights, n = 10 )
		, formats = get.light.formats()
)
doc = addLineBreak( doc )

# simple addTable with row merging on 'day' column
doc = addHeader(doc, "Dataset last lines", 2);
sampledata = tail( measured.weights[ order(measured.weights$day), ], n = 20 )
doc = addTable( doc
		, data = sampledata
		, span.columns = c("day")
		, col.colors = list( "gender" = ifelse( sampledata$gender == "male" , "#00c2ff", "#ffcdd1") )
)
doc = addPageBreak( doc )

# customized table
doc = addHeader(doc, "Dataset summary", 1);
doc = addTable( doc
		, data = weights.summary
		, header.labels = list("id" = "Subject Identifier", "avg.weight" = "Average Weight"
				, "regularity" = "Regularity", "visit.n" = "Number of visits", "last.day" = "Last visit"
		) # columns labels to display
		, grouped.cols=list( "id" = "id"
				, "Summary" = c( "avg.weight",  "regularity", "visit.n", "last.day" ) 
		) # grouped headers line to add before headers line
		, col.types = list( "id" = "character", "avg.weight" = "double"
				, "regularity" = "percent", "visit.n" = "integer"
				, "last.day" = "date" ) # reporting types of each columns 
		, col.fontcolors = list( "regularity" = ifelse( weights.summary$regularity < 0.5 , "gray", "blue") ) # customized font colors for column "regularity"
)

doc = addPageBreak( doc )
doc = addHeader(doc, "Graphics", 1);
doc = addPlot( doc
		, fun = boxplot
		, formula = measured.weights$weight ~ measured.weights$id
		, xlab = "Subjects", ylab = "Weights" , col = "#0088cc"
		, legend = "My boxplot"
)

require( ggplot2 )
my.ggplot = qplot(day, weight, data = measured.weights, geom = "line", color = id )
doc = addPlot( doc
		, fun = print
		, x = my.ggplot
		, legend = "ggplot example"
		, width = 9, height = 7
)
writeDoc( doc, docx.file )

if( interactive() ) {
	out = readline( "Open the docx file (y/n)? " )
	if( out == "y" ) browseURL( file.path(getwd(), docx.file ) )
}
