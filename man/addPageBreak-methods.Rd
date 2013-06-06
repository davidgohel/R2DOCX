\name{addPageBreak-methods}
\docType{methods}
\alias{addPageBreak-methods}
\title{Insert a page break into the docx document}
\description{
Insert a page break into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{addPageBreak}{Docx}(x)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
}


\examples{
# example 1
doc <- new("Docx", title = "My example" )
doc <- addPageBreak(doc)
doc <- addParagraph( doc, value = "Hello", stylename = "Normal")
writeDoc( doc, "example.docx" )



}
\keyword{methods}
