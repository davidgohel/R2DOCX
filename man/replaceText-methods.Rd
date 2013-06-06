\name{replaceText-methods}
\docType{methods}
\alias{replaceText-methods}
\title{Replace text into the docx document}
\description{
Replace text into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{replaceText}{Docx,character,character}(x,pattern,replacement)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{pattern}{text to replace (character vector of length 1)}
  \item{replacement}{value used as replacement(character vector of length 1)}
}
\details{
This function scan the entire document and then replace values. 

Goal of this function is to enable title, date, or any other 'field'  
replacements into templates, best is to call the function before 
any element insertion when the doc is still light in memory.

}
\examples{
# first we need to create a doc where we put texts to replace later.
doc <- new("Docx", title = "My example" )
doc <- addParagraph( doc, value = "donkeys eats grass", stylename = "Normal" )
writeDoc( doc, "template.docx" )

# now we will read the "template.docx" and replace 'donkeys' by 'cows'
anotherdoc <- new("Docx", title = "My example", basefile = "template.docx" )
anotherdoc <- replaceText( anotherdoc, "donkeys", "cows")
writeDoc( anotherdoc, "output.docx" )
}
\keyword{methods}
