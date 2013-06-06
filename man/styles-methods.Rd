\name{styles-methods}
\docType{methods}
\alias{styles-methods}
\title{Get existing styles}
\description{
Get styles names that exist into the template (base document).
}
\usage{
  \S4method{styles}{Docx}(x)
}
\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
}
\details{
\code{styles} is a key function in the R2DOCX package.\cr
It return existing styles in the base document used as a template.\cr
}
\examples{
doc = new( "Docx" , "example")
styles(doc)
}
\keyword{methods}
