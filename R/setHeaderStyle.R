# TODO: Documentation de setHeaderStyle.R
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 22 avr. 2013
# Version: 0.1
###############################################################################


setMethod("setHeaderStyle", "Docx", function( x, stylenames ) {
			if( !all( is.element( stylenames, styles( x ) ) ) ){
				stop("Some of the stylenames are not in available styles (run styles on your object to list available styles.")
			}
			x@header.styles = stylenames
			x
		})

