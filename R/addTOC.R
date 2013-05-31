# Method Docx.addTOC 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################
#TODO : check stylename is existing
setMethod("addTOC", "Docx", function(x, stylename, ... ) {
			if( missing( stylename ) ){
				.jcall( x@obj, "V", "addTableOfContents" )
			} else {
				if( !is.element( stylename , styles( x ) ) ){
					stop(paste("Style {", stylename, "} does not exists.", sep = "") )
				}
				else .jcall( x@obj, "V", "addTableOfContents", stylename )
			}
			x
	})


