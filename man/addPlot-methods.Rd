\name{addPlot-methods}
\docType{methods}
\alias{addPlot-methods}
\title{Insert a plot into the docx document}
\description{
Insert a plot into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{addPlot}{Docx,function,numeric,numeric,character,character,ANY}(x, fun, width, height, legend, stylename,...)
}
\arguments{
  \item{doc}{the \code{\linkS4class{Docx}} to use}
  \item{fun}{plot function}
  \item{width}{plot width in inches.}
  \item{height}{plot height in inches.}
  \item{legend}{character string, legend to add below the plot}
  \item{stylename}{name of the style to apply to the legend text}
  \item{...}{arguments for \code{fun}.}
}
\details{
This function work with base R graphics. Plot parameters are specified with 
the \code{...} argument. However, the most convenient usage is to wrap 
the plot code into a function whose parameters will be specified as '...'.

If you want to add ggplot2 or lattice plot, use 
\code{print} function. See examples for more details.
}
\examples{
require(ggplot2)
# Create a new document 
doc <- new("Docx", title = "title" )
# Add a base plot
doc = addPlot( doc
		, fun = plot
		, width = 7, height = 6, legend="base plot" 
		, stylename = "PlotReference"
		, x = rnorm( 100 )
		, y = rnorm (100 )
		, main = "base plot main title"
	)
writeDoc( doc, "Hello_Word.docx" )


## Add a ggplot2
# method 1 :
doc <- new("Docx", title = "title" )
doc = addPlot( doc
		, function(){
			print( qplot(Sepal.Length, Petal.Length, data = iris, color = Species, size = Petal.Width, alpha = I(0.7)) )
		}
		, width = 7, height = 6, legend="graph example" 
		, stylename = "PlotReference"
	)
writeDoc( doc, "Hello_Word.docx" )

# method 2 :
myplot = qplot(Sepal.Length, Petal.Length, data = iris, color = Species, size = Petal.Width, alpha = I(0.7))
doc = addPlot( doc
		, print
		, width = 7, height = 6, legend="graph example" 
		, stylename = "PlotReference"
		, x = myplot #this argument MUST be named, print is waiting arg 'x'
	)
\dontrun{
writeDoc( doc, "Hello_Word.docx" )
}
}
\keyword{methods}
