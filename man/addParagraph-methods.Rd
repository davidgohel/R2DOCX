\name{addParagraph-methods}
\docType{methods}
\alias{addParagraph-methods}

\title{Insert a paragraph of text into the docx document}

\description{
Insert a paragraph of text into a \code{\linkS4class{Docx} object}.
}

\usage{
  \S4method{addParagraph}{Docx, character, character, list, list}(x, value, par.style, replacements, replacement.styles)
}


\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{value}{a character (vector of length 1) containing text to add.}
  \item{par.style}{value of the named style in the document.}
  \item{replacements}{a named list whose elements define replacements values.}
  \item{replacement.styles}{a named list specifying formats to apply to replacements.}
}

\details{

\itemize{
\item \code{value} : {text to add, a character vector of length 1.\cr
		Example : \code{value="Hello world"}\cr
		If replacements have to be done, keywords that identify text to be replaced must be contains into square brackets [].\cr
		Example : \code{value="At the [ALPHA] significance level, one [DECISION] the null hypothesis."}\cr
		}

\item \code{par.style} : {Stylename to use in the document, a character vector of length 1.\cr
		Example : \code{par.style="Normal"}\cr
		Stylename must exist in the base document provided as a template to the \code{\linkS4class{Docx}} object.\cr
		Use \code{styles(Docx)} to see available styles in the base document template.\cr
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
		list( animal= new ("textProperties", "color"="red" ) , food= new ("textProperties", "color"="blue" ) )
		}\cr
		}

}


}



\examples{
doc <- new("Docx", title = "My example" )

x = "[animal] eats [food]."

repl.styles = list(
		  animal= new ("textProperties", "font-size" = "12px", "font-family"="Courier New", "color"="red" )
		, food= new ("textProperties", "font-size" = "12px", "font-family"="Courier New", "color"="blue" )
)
repl = list( animal = "donkey" , food = "grass" )

doc <- addParagraph( doc, value = x, par.style = "Normal"
		, replacements = repl
		, replacement.styles = repl.styles
)

writeDoc( doc, "Paragraph_example.docx" )

}