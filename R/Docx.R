# Docx class definition
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setClass("Docx"
	, representation(obj = "ANY"
		, title = "character"
		, basefile = "character"
		, styles = "character"
		, header.styles = "character"
		))

setMethod("show", "Docx",
		function(object){
			cat("Document title : ")
			cat(object@title)
			cat("\n")
			cat(paste( "Original document : '", object@basefile, "'\n",sep = "" ) )
})
