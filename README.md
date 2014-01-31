# NOTICE 28 January 2014

R2DOCX development activity has shifted to [`ReporteRs`](https://github.com/davidgohel/ReporteRs).

**ReporteRs** generate Word, PowerPoint and html files. R2DOCX functions have been improved in ReporteRs package. 

R2DOCX
======
R2DOCX is a R package for creating Word docx files (Microsoft Open XML).

Here is the [documentation](http://davidgohel.github.io/R2DOCX/index.html).

Features
--------
* Create docx files with only a few lines of R code. 
* Add tables, plots, texts or tables of contents into a word document.
* Simple R functions are available for customizing formatting properties of R output.

Installation
------------
From an R console (R >= 3.0):

    install.packages("devtools")
    devtools::install_github('R2DOC', 'davidgohel')
    devtools::install_github('R2DOCX', 'davidgohel')

**Dependencies**

    R packages : rJava
	Java (it has been tested with java version >= 1.6)
	
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


License
-------
The R2DOCX package is licensed under the GPLv3. For additional details, see files:
* COPYING - R2DOCX package license (GPLv3)
* NOTICE - Copyright notices for additional included software

