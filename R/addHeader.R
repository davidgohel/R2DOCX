# TODO: Documentation de addHeader.R
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 10 avr. 2013
# Version: 0.1
###############################################################################

setMethod("addHeader", "Docx", function( x, value, level ) {
			if( length( x@header.styles ) == 0 ){
				stop("You must defined header styles via setHeaderStyle first.")				
			}
			if( length( x@header.styles ) < level ){
				stop("level = ", level, ". You defined only ", length( x@header.styles ), " styles.")				
			}
			
			x <- addParagraph(x, value, x@header.styles[level] );
			x
		})


