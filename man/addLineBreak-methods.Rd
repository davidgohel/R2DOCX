\name{addLineBreak-methods}
\alias{addLineBreak}
\docType{methods}
\title{Insert a page break into the docx document}
\description{
Insert a line break into the \code{\linkS4class{Docx}} object.
}
\usage{
  addLineBreak(x)
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
