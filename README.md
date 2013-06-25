R2DOCX
======
R2DOCX is a R package for creating Word docx files (Microsoft Open XML).
Features
--------
* Create docx files with only a few lines of R code. 
* Add tables, plots, texts or tables of contents into a word document.
* Simple R functions are available for customizing formatting properties of R output.

Installation
------------
From an R console (R >= 3.0):

    devtools::install_github('R2DOC', 'davidgohel')
    devtools::install_github('R2DOCX', 'davidgohel')

Getting Started
---------------

    library(R2DOCX)
    demo(R2DOCX) #run a complete and detailed example
    
Creating Microsoft Word documents is easy and can be done in few lines of R code. Below a commented R script:

    # Here we define word document filename to write
    docx.file <- "document.docx"
    # Creation of doc, a Docx object
    doc <- new("Docx", title = "My example" )
    # add into doc first 10 lines of iris
    doc = addTable( doc, iris[1:10,] )
    # add text with stylename ‘Normal’ into doc 
    doc <- addParagraph( doc, value = "Hello World!", stylename = "Normal" )
    # add a plot into doc 
    doc = addPlot( doc, function() plot( rnorm(10), rnorm(10) )
        , width = 10, height = 8
    )
    # write the doc 
    writeDoc( doc, docx.file )


Summary
-------
When creating a Docx object, a base file is used. This file is the original docx document 
that will be completed with R ouputs. This file will be copied "in memory". 

Available styles will be those available in the base file. Tables and paragraphs formats 
can be customize with dedicated functions.

To create a docx document from R, use:

    doc <- new("Docx", title = "My example" )

To add R outputs (or replace text) in a docx document, use one of these functions:
* 'addPlot' for adding plots.
* 'addParagraph' for adding paragraphs of text.
* 'addTable' for adding tables.
* 'addHeader' for adding headers.
* 'addTOC' for adding tables of contents.
* 'addPageBreak' for adding a page break.
* 'addLineBreak' for adding a line break.
* 'replaceText' for text replacement.

To write the docx document on to the disk:

    writeDoc( doc, "my_example.docx" )

### Note
If you see that error:

    Error in addHeader(...
        You must defined header styles via setHeaderStyle first.

it means you have to use function 'setHeaderStyle' to indicate which available 
styles are meant to be used as header styles.

    doc <- new("Docx", title = "My example" )
    styles( doc )
    # [1] "Normal"                  "Titre1"                  "Titre2"                 
    # [4] "Titre3"                  "Titre4"                  "Titre5"                 
    # [7] "Titre6"                  "Titre7"                  "Titre8"                 
    #[10] "Titre9"                  "Policepardfaut"          "TableauNormal"          
    # ...
    doc <- setHeaderStyle(doc, stylenames = c("Titre1", "Titre2", "Titre3", "Titre4", "Titre5", "Titre6", "Titre7", "Titre8", "Titre9" ) ) 
    doc = addHeader( doc, "title 1", 1 )

Output tables
-------------
Simple table output can be done with function 'addTable':

    doc <- new("Docx", title = "My example" )
    doc = addTable( doc, data = iris )

Data argument must be a data.frame object. 

Arguments are available for specifying more complex formats.

### Table model ###

    --------------------+---------------------+
      GROUPED HEADER 1  |   GROUPED HEADER 1  |
    ---------+----------+----------+----------+
    HEADER 1 | HEADER 2 | HEADER 3 | HEADER 4 |
    ---------+----------+----------+----------+
    data[1,1]| data[1,2]| data[1,3]| data[1,2]|
    ---------+----------+----------+----------+
    data[2,1]| data[2,2]| data[2,3]| data[2,2]|
    ---------+----------+----------+----------+
          ...|       ...|       ...|       ...|
    ---------+----------+----------+----------+

**Features**
* Grouped headers are an optional header row. Its purpose is to gather headers under a label. All grouped headers share the same formatting properties.
* Headers are columns main headers, by default, columns names are used as headers labels. All headers share the same formatting properties.
* Data cells will be formatted conditionally to the declared column type. Default values are ‘integer’ for ‘integer’ data, ‘double’ for ‘double’ data, ‘character’ for ‘character’ and ‘factor’ data, ‘date’ for ‘Date’ data. Available types are ‘integer’, ‘double’, ‘character’, ‘percent’, ‘date’, ‘datetime’. 
* Optionally, data cells can be colored (cell background and or font color).
* Optionally, data cells can be merged when values are repeated inside a column.

### addTable tutorial ###

**Headers labels**

    desc( iris )
    #'data.frame':   150 obs. of  5 variables:
    # $ Sepal.Length: num  5.1 4.9 4.7 4.6 5 5.4 4.6 5 4.4 4.9 ...
    # $ Sepal.Width : num  3.5 3 3.2 3.1 3.6 3.9 3.4 3.4 2.9 3.1 ...
    # $ Petal.Length: num  1.4 1.4 1.3 1.5 1.4 1.7 1.4 1.5 1.4 1.5 ...
    # $ Petal.Width : num  0.2 0.2 0.2 0.2 0.2 0.4 0.3 0.2 0.2 0.1 ...
    # $ Species
    doc <- addTable( doc , iris
        , header.labels = list("Sepal.Length" = " Sepal Length"
    	, "Sepal.Width" = " Sepal Width"
    		, "Petal.Length" = " Petal Length", "Petal.Width" = " Petal Width"
    		, "Species" = "Species"
    		)
    	)

**Grouped headers labels**

    doc <- addTable( doc
        , iris
        , grouped.cols=list( "Sepal" = c( "Sepal.Length", "Sepal.Width" )
            , "Petal" = c("Petal.Length", "Petal.Width")
            , "Species" = c("Species") 
		)
	)

Argument is a named list. Names are labels to display, list elements are character vector 
containing data column names to gather. All columns names must be in exactly one element. 
Order of appearance in elements must be the order of columns names in the dataset. 
E.g. if columns are ‘col1’, ‘col2’, ‘col3’, argument cannot be: 

    list (‘colgroup1’=’col2’, ‘colgroup2’=c(‘col1’, ’col3’))

To leave blank headers labels, name the element the same that the column name:

    "Species" = "Species"


## Output plots
Function 'addPlot' will capture plots into png files and will insert them into the output 
document with a legend text below graphics. Arguments ‘legend’ and ‘stylename’ are detailed 
in the chapter “styles”. 

Its main argument is a function. The plotting is done by this function. If this function 
requires parameters, add them as parameters to the ‘addPlot’ function (the … argument).

    doc <- addPlot( doc , plot
        , x = rnorm( 10 ), y = rnorm( 10 ), main = "main title"
        , legend = "graph example", stylename="PlotReference" 
    )

If manipulating lattice or ggplot2 plot objects, use print on the object in the function (otherwise lattice or ggplot2 objects won’t be written). 

    doc <- addPlot( doc , print
        , x = qplot(Sepal.Length, Petal.Length, data = iris
        , color = Species, size = Petal.Width )
    , legend = "graph example", stylename="PlotReference" 
    )

To combine multiple plots, simply create a function to execute all plotting instructions:

    doc <- addPlot( doc , function( ) {
        print( qplot(Sepal.Length, Petal.Length
            , data = iris, color = Species, size = Petal.Width ) )
        plot( x = rnorm( 10 ), y = rnorm( 10 ), main = "main title" )
        }
    , legend = "graph example", stylename="PlotReference" 
    )
    
## Output texts

Simple texts output can be done with function ‘addParagraph’:

    doc <- new("Docx", title = "My example" )
    doc = addParagraph( doc, "Hello World!", stylename = "Normal" )

The function can also be used to execute text replacement and conditional formatting. 
If replacements have to be done, keywords that identify text to be replaced must be contains 
into square brackets [].

**Detailed example**

    # template text
    x = "[animal] eats [food]."

    # define formatting properties for replacement text – see chapter "styles"
    textProp = textProperties( color = "blue" )
    replacement.styles = list( animal= textProp, food= textProp )
    # define replacement text
    replacements = list( animal = "donkey", food = "grass" )
    
    doc <- addParagraph( doc, value = x, stylename = "Normal"
        , replacements = replacements
    	, replacement.styles = replacement.styles
    	)


## text replacement

Funtion 'replaceText' can be executed to replace texts into a docx document. 
E.g. if we want to replace a keyword ‘AUTHOR’ by "John Doe", simply run:

    doc = replaceText( doc, "AUTHOR", "John Doe")

License
-------
The R2DOCX package is licensed under the GPLv3. For additional details, see files:
* COPYING - R2DOCX package license (GPLv3)
* NOTICE - Copyright notices for additional included software

