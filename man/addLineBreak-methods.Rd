\name{addLineBreak-methods}
\docType{methods}
\alias{addLineBreak-methods}
\title{Insert a page break into the docx document}
\description{
Insert a line break into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{addLineBreak}{Docx}(x)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
}


\examples{
# example 1
doc <- new("Docx", title = "My example" )
doc <- addLineBreak(doc)
doc <- addLineBreak(doc)
doc <- addParagraph( doc, value = "Hello", stylename = "Normal")
writeDoc( doc, "example.docx" )



}
\keyword{methods}
