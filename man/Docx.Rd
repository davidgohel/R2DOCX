\name{Docx-class}
\Rdversion{1.1}
\docType{class}
\alias{Docx-class}

\title{Microsoft Word document representation - class \code{"Docx"}}
\description{
Objects from the class \code{"Docx"} represent Microsoft Word document. 
All corresponding methods are used to send R output into this object.
}

\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("Docx", title="Title of the document", basefile)}.
}
\section{Slots}{
  \describe{
    \item{\code{title}}{Object of class \code{"character"}:
      title of the document (in the doc properties). }
    \item{\code{basefile}:}{Object of class \code{"character"} 
      which represents the filename of the docx file used as a template. }
  }
}

\section{Methods}{
  \describe{
    \item{addPlot}{\code{signature(x = "Docx", ...)}: Add a graph.}
    \item{addParagraph}{\code{signature(x = "Docx", ...)}: Add a paragraph.}
    \item{addTable}{\code{signature(x = "Docx", ...)}: Add a table.}
    \item{addTOC}{\code{signature(x = "Docx", ...)}: Add a table of contents.}
    \item{styles}{\code{signature(x = "Docx", ...)}: get available styles.}
    \item{replaceText}{\code{signature(x = "Docx", ...)}: replace text.}
    \item{writeDoc}{\code{signature(x = "Docx", ...)}: write the docx object in a file.}
	 }
}

\references{
Wikipedia: Office Open XML\cr
\url{http://en.wikipedia.org/wiki/Office_Open_XML}
}

\author{
David Gohel
}
\note{
\pkg{R2DOCX} supports Word 2007-2013 (*.docx) file formats.\cr
Document are manipulated in-memory ; a \code{Docx}'s document is not written to the disk unless the \code{\link[=writeDoc-methods]{writeDoc}} method has been called on the object. 
}
\examples{
# Create a new document 
doc <- new("Docx", title = "title" )
# Add "Hello World" into the document
doc <- addParagraph(doc, "Hello Word!", "Normal");
# writes document into file "Hello_Word.docx"

writeDoc( itsDoc, "Hello_Word.docx" )
}
\concept{Microsoft Word}
\concept{docx}
\keyword{package}
\keyword{print}
