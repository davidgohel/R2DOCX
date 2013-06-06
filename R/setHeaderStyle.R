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

#setMethod("addStyle", "Docx", function( x, stylename, color = "black", font.size = 10
#	, font.weight = "normal", font.style = "normal"
#	, font.family = "Arial"
#	, text.align = "left", padding.bottom = 1
#	, padding.top = 1, padding.left = 1
#	, padding.right = 1, padding ){
#
#	if( !missing( padding ) ){
#		if( is.numeric( padding ) ) {
#			if( as.integer( padding ) < 0 || !is.finite( as.integer( padding ) ) ) stop("invalid padding : ", padding )
#			padding.bottom = padding.top = padding.right = padding.left = as.integer( padding )
#		} else {
#			stop("padding must be a integer value ( >=0 ).")
#		}
#	}
#	#TODO:add controls here...
#	.jcall( x@obj, "V", "createStyle"
#			, stylename
#			, color
#			, as.integer( font.size )
#			, as.integer(font.weight=="bold"), as.integer(font.weight=="italic")
#			, font.family
#			, text.align
#			, as.integer( padding.top ), as.integer( padding.bottom ), as.integer( padding.left ), as.integer( padding.right )
#	)
#
#	x
#})

