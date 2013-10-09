\name{addImage-methods}
\docType{methods}
\alias{addImage-methods}
\title{Insert an external image into the docx document}
\description{
Insert an external image into the \code{\linkS4class{Docx}} object.
}
\usage{
  \S4method{addImage}{Docx,character,character}(x, filename, bookmark)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{filename}{a character value ; image filename}
  \item{bookmark}{a character value ; id of the Word bookmark to replace by the image. optional}
}
\examples{
\dontrun{
doc <- new("Docx", title = "My example" )
doc = addImage( doc, "D:/img/image.png")
}
}
\keyword{methods}
