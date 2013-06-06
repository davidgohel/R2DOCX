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
  \item{span.columns}{a character vector of columns names where row merging is to be done.}
  \item{col.types}{Defines the formating style of columns}
  \item{col.colors}{Defines the background color of cells}
  \item{col.fontcolors}{Defines the font color of cells}
}

\details{
\preformatted{
--------------------+---------------------+
  GROUPED HEADER 1  |   GROUPED HEADER 1  |
---------+----------+----------+----------+
HEADER 1 | HEADER 2 | HEADER 3 | HEADER 4 |
---------+----------+----------+----------+
data[1,1]| data[1,2]| data[1,3]| data[1,2]|
---------+----------+----------+----------+
data[2,1]| data[2,2]| data[2,3]| data[2,2]|
---------+----------+----------+----------+
      ...|       ...|       ...|       ...|
---------+----------+----------+----------+
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


\item \code{span.columns} : {a character vector  
		specifying columns names where row merging should 
		be done (only if data[p,j]==data[p-1,j] )\cr
		Example : \code{data = iris, span.columns = "Species"} mean to merge 
		identical values for the column "Species" of the iris dataset.\cr
		Default is character(0)
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
data( measured.weights )
data( weights.summary )

# Word document to write
docx.file <- "document.docx"

# create document
doc <- new("Docx", title = "My example" )

# add a title
doc <- addHeader( doc, "Table example", 1 )

# add the first 5 lines of measured.weights in the docx
doc <- addTable( doc, head( measured.weights, n = 5 ) )

# demo span.columns
doc <- addHeader( doc, "Another table example", 1 )
doc <- addTable( doc, iris[ 46:55,], span.columns = "Species" )


# add another title
doc <- addHeader( doc, "And a last one", 1 )

# add and format weights.summary in the docx
doc <- addTable( doc
		, data = weights.summary
		, header.labels = list("id" = "Subject Identifier", "avg.weight" = "Average Weight"
				, "regularity" = "Regularity", "visit.n" = "Number of visits", "last.day" = "Last visit")
		, grouped.cols=list( "id" = "id"
				, "Summary" = c( "avg.weight",  "regularity", "visit.n", "last.day" ) )
		, col.types = list( "id" = "character", "avg.weight" = "double"
				, "regularity" = "percent", "visit.n" = "integer"
				, "last.day" = "date" ) 
		, col.fontcolors = list( "regularity" = ifelse( weights.summary$regularity < 0.5 , "gray", "blue") )
)

# write the docx onto the disk
writeDoc( doc, docx.file )

}