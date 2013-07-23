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

    install.packages("devtools")
    devtools::install_github('R2DOC', 'davidgohel')
    devtools::install_github('R2DOCX', 'davidgohel')

**Mac users**

    Sys.setenv(NOAWT=1) #prevents usage of awt.[http://cran.r-project.org/doc/manuals/r-devel/R-admin.html#Java-_0028OS-X_0029]
	install.packages("devtools")
    devtools::install_github('R2DOC', 'davidgohel')
    devtools::install_github('R2DOCX', 'davidgohel')

**Dependencies**

    R packages : rJava
	Java (it has been tested with java version >= 1.6)
	
Documentation
-------------
[R2DOCX](http://davidgohel.github.io/R2DOCX/index.html)

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

