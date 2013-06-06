# Method Docx.writeDoc 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("writeDoc", "Docx", function(x, file, ...) {
			if( !tryCatch( {
					cat("", file = file)
					T
					}, error = function( e ) F, warning = function ( e ) F , finally = T) )
				stop("writeDoc: Cannot write to ", file)
			
			.jcall( x@obj , "V", "writeDocxToStream", file )
		})



