\name{addTOC-methods}
\docType{methods}
\alias{addTOC-methods}
\title{Insert a plot into the docx document}
\description{
Insert a plot into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{addTOC}{Docx,character}(x, stylename)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{stylename}{name of the style used to construct the TOC}
}
\details{
This function add a table of contents in a doc object.\cr
It can also add custom table of contents, for example a table of content for the added plots.

\itemize{
\item \code{stylename} : {Stylename in the document that will 
			be used to build entries of the TOC (a character vector of length 1).\cr
		}

}
}
\examples{
# example 1
doc <- new("Docx", title = "My example" )
doc <- addTOC(doc)
# add headers and content here
writeDoc( doc, "toc_example.docx" )




# example 2
require(ggplot2)
doc <- new("Docx", title = "My example" )
# Because we will use "PlotReference" as legend in plot 
# later, addTOC will use this stylename to define 
# entries in the generated TOC 
doc <- addTOC(doc, stylename = "PlotReference")

doc = addPlot( doc
		, fun = plot
		, x = rnorm( 100 )
		, y = rnorm (100 )
		, main = "base plot main title"
		, legend="graph example 1", stylename = "PlotReference"
	)

doc = addTable( doc, iris )
doc = addParagraph( doc, value="table 1", stylename = "PlotReference" )

doc = addPlot( doc
		, fun = plot
		, x = rnorm( 100 )
		, y = rnorm (100 )
		, main = "base plot main title"
		, legend="graph example 2", stylename = "PlotReference"
	)


writeDoc( doc, "toc_example.docx" )

}
\keyword{methods}
