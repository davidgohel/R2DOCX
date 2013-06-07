\name{writeDoc-methods}
\docType{methods}
\alias{writeDoc-methods}
\title{Write the docx document in a docx file.}
\description{
Write the \code{\linkS4class{Docx}} object in a '.docx' file.
}
\usage{
  \S4method{writeDoc}{Docx,character}(x,file)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{file}{single character value, name of the file to write.}
}
\details{
The function cannot write a document if it already exists and is opened.
}
\examples{
doc <- new("Docx", title = "My example" )
doc <- addParagraph( doc, value = "Hello world", stylename = "Normal" )
writeDoc( doc, "template.docx" )

}
\keyword{methods}
