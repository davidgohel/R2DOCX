\name{addParagraph-methods}
\docType{methods}
\alias{addParagraph-methods}

\title{Insert a paragraph of text into the docx document}

\description{
Insert a paragraph of text into a \code{\linkS4class{Docx} object}.
}

\usage{
  \S4method{addParagraph}{Docx, character, character, list, list, character}(x, value, stylename, replacements, replacement.styles, bookmark)
}


\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{value}{character vector containing texts to add.}
  \item{stylename}{value of the named style in the docx document.}
  \item{replacements}{optional, a named list whose elements define replacements values.}
  \item{replacement.styles}{optional, a named list specifying formats to apply to replacements.}
  \item{bookmark}{a character value ; id of the Word bookmark to replace by the table. optional}
}

\details{

\itemize{
\item \code{value} : {text(s) to add, a character vector.\cr
		Example : \code{value="Hello world"}\cr
		If replacements have to be done, keywords that identify text to be replaced must be contained into square brackets [].\cr
		Example : \code{value="At the [ALPHA] significance level, one [DECISION] the null hypothesis."}\cr
		}


\item \code{replacements} : {a name list, names are keywords (of the \code{value}
		, contained into square brackets) and values are replacements values\cr
		Example : \code{value = "[animal] eats [food]."}\cr\code{, replacements = 
		list( animal = "donkey" , food = "grass" )}\cr
		}

\item \code{replacement.styles} : {a name list, names are keywords (of the \code{value}
		, contained into square brackets) and values are replacements styles values\cr
		Example : \code{value = "[animal] eats [food]."}\cr\code{, replacements = 
		list( animal = "donkey" , food = "grass" )}\cr\code{, replacement.styles = 
		list( animal= textProperties(color="red" ) , food= textProperties(color="blue" ) )
		}\cr
		}

\item \code{bookmark} : {a character vector specifying bookmark id to replace by the table. 
		If provided, table will replace the paragraph that contains the bookmark. 
		If not provided, table will be added at the end of the document.
		}

}


}



\examples{
doc <- new("Docx", title = "My example" )

doc <- addParagraph( doc, value = "Hello", stylename = "Normal")




x = c( "[animal] eat [food].", "tigers eat [animal]." )

repl.styles = list(
		  animal= textProperties( font.size = 12, font.family="Courier New", color="red" )
		, food= textProperties( font.size = 12, font.family="Courier New", color="blue" )
)
repl = list( animal = "buffalos" , food = "grass" )

doc <- addParagraph( doc, value = x, stylename = "Normal"
		, replacements = repl
		, replacement.styles = repl.styles
)
doc <- addParagraph( doc, value = letters[1:10], stylename = "BulletList")

\dontrun{
writeDoc( doc, "Paragraph_example.docx" )
}
}