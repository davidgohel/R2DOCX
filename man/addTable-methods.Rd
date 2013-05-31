\name{addTable-methods}
\docType{methods}
\alias{addTable-methods}

\title{Insert a table into the docx document}

\description{
Insert a table into a \code{\linkS4class{Docx} object}.
}

\usage{
  \S4method{addTable}{Docx, data.frame, tableProperties, list, list, integer, character, list, list}(x, data, formats, header.labels, grouped.cols, span.first.columns=0, col.types, col.colors, col.fontcolors)
}


\arguments{
  \item{x}{the \code{\linkS4class{Docx}} to use}
  \item{data}{\code{data.frame} to add}
  \item{formats}{a \code{\linkS4class{tableProperties} object} that specify styles to use to format the table.}
  \item{header.labels}{a named list whose elements define labels to display in table headers instead of colnames.}
  \item{grouped.cols}{a named list whose elements define which columns are to be grouped in an upper header row (grouped header).}
  \item{span.first.columns}{Specify the number of columns where row merging is to be done.}
  \item{col.types}{Defines the formating style of columns}
  \item{col.colors}{Defines the background color of cells}
  \item{col.fontcolors}{Defines the font color of cells}
}

\details{
\preformatted{
    -------------------------------------------
      GROUPED HEADER 1  |   GROUPED HEADER 1  |
    -------------------------------------------
    HEADER 1 | HEADER 2 | HEADER 3 | HEADER 4 |
    -------------------------------------------
    data[1,1]| data[1,2]| data[1,3]| data[1,2]|
    data[2,1]| data[2,2]| data[2,3]| data[2,2]|
          ...|       ...|       ...|       ...|
    -------------------------------------------
}


\itemize{
\item \code{header.labels} : {a name list, names are \code{data} 
		column names and values are labels to use in HEADER 
		cells(character vector of lenght 1)\cr
		Example : \code{list("col1"="header 1","col2"="header 2")}\cr
		If missing, headers will be filled with \code{data} 
		column names
		}

\item \code{grouped.cols}{a name list, names are \code{data} 
		grouped headers labels and value are \code{data} 
		column names gathered under the grouped headers labels\cr
		Example : \code{list("grouped header 1"=c( "col1", "col2"),"grouped header 2"=c( "col3", "col4"))}
		\cr
		If missing, grouped headers won't be added to the table.\cr
		If used, columns order must be respected and all \code{data} 
		column names must belong to an element of the list:\cr
		Example (WRONG) : \code{list("grouped header 1"=c( "col3", "col2"),"grouped header 2"=c( "col1", "col4"))}\cr
		Example (WRONG) : \code{list("grouped header 1"=c( "col1", "col2"),"grouped header 2"=c("col4"))}
	}


\item \code{span.first.columns} : {an integer 
		specifying number of columns where row merging should 
		be done (only if data[p,j]==data[p-1,j] )\cr
		Example : \code{span.first.columns = 2} mean to merge 
		identical values for the first 2 columns of the table.\cr
		Default is 0. Value must be less than the number of data columns.
		}


\item \code{col.types} : {a name list, names are \code{data} 
		column names and values are specifying how to format data in the table\cr 
		Possible values are : \emph{"character"}, \emph{"integer"}, \emph{"double"}, \emph{"percent"}, \emph{"date"}, \emph{"datetime}"\cr 
		Example : \code{list("col1"="character","col2"="double")}\cr
		If missing, factor and character will be formated as character, integer as integer and numeric as double.
		}


\item \code{col.colors} : {a name list, names are \code{data} 
		column names and values are character vectors specifying cells background colors\cr
		Each element of the list is a vector of length \code{nrow(data)}.\cr
		Example : \code{list("col1"=c("#AAEE00", "#AAEE00", "#000000", ...) , "col2"=c("#000000", "#AAEE00", "#AAEE00", ...) )}\cr
		This argument is optional.
		}

\item \code{col.fontcolors} : {a name list, names are \code{data} 
		column names and values are character vectors specifying cells font colors\cr
		Each element of the list is a vector of length \code{nrow(data)}.\cr
		Example : \code{list("col1"=c("#AAEE00", "#AAEE00", "#000000", ...) , "col2"=c("#000000", "#AAEE00", "#AAEE00", ...) )}\cr
		This argument is optional.
		}




}


}



\examples{
require(lattice)

# simple example
doc <- new("Docx", title = "My example" )
doc <- addTable( doc, iris )
writeDoc( doc, "table_example.docx" )


# more sophisticated example
data = barley[, c("variety", "site", "year", "yield" )]
data$PC = runif( nrow( data ), 0, 1 )# create dummy percent

doc <- new("Docx", title = "My example" )
doc <- addTable( doc
	, data
	, header.labels = list("variety" = "Variety", "site" = "Site"
		, "year" = "Year", "yield" = "Yield", "PC" = "Percent column")
	, grouped.cols=list( "Types" = c( "variety", "site" )
		, "year" = "year"
		, "Mesures" = c("yield", "PC") )
	, span.first.columns = 1
	, col.types = list("variety" = "character", "site" = "character"
		, "year" = "integer", "yield" = "double", "PC"="percent" ) 
	, col.fontcolors = list( "PC" = ifelse( data$PC < 0.05 , "#FF0000", "#000000") )
)
writeDoc( doc, "another_table_example.docx" )

}