# Method Docx.addParagraph 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("addParagraph", "Docx", function(x, value, par.style, replacements, replacement.styles, ... ) {
	if( missing( par.style )) {
		stop("argument 'par.style' is missing")
	} else if( !is.element( par.style , styles( x ) ) ){
		stop(paste("Style {", par.style, "} does not exists.", sep = "") )
	}
	if( !missing( replacement.styles ) ){
		if( missing( replacements ) ) stop("argument 'replacements' is missing while argument 'replacement.styles' has been supplied")
		if( !all( sapply( replacement.styles, is, "textProperties" ) ) ) stop("argument 'replacement.styles' must be a list of 'textProperties' objects.")
	} else if( !missing( replacements ) ){
		stop("argument 'replacement.styles' is missing while argument 'replacement' has been supplied")
	}
	
	.jcall( x@obj, "V", "initParagraphTemplate" )
	if( !missing( replacements ) ){
		pattern <- "\\[{1}[a-zA-z]+\\]{1}"#any '[a least a letter and no space]'
		m = unlist( regmatches(value, gregexpr( pattern, value ) ) )
		
		if( !all( is.element( paste( "[", names(replacements), "]", sep = "" ), m )  ) ){
			stop("Some of the replacement cannot be found in the parameter 'value'.")
		}
		if( !all( is.element( m, paste( "[", names(replacements), "]", sep = "" ) )  ) ){
			warning("Replacements seem to be missing in the parameter 'replacements'.")
		}

		# loop over each replacement
		for( j in names( replacements ) ){
			fs = replacement.styles[[j]]@properties["font-size"]
			fs = as.integer( gsub( "px", "", fs ) )
			
			.jcall( x@obj, "V", "addParagraphTemplateReplacement"
					, paste("[", j, "]", sep = "" )#keys
					, replacements[[j]] #value
					, fs #font-size
					, replacement.styles[[j]]@properties["font-weight"]
					, replacement.styles[[j]]@properties["font-style"]
					, replacement.styles[[j]]@properties["color"]
					, replacement.styles[[j]]@properties["font-family"]
			)
		}
	}

	.jcall( x@obj, "V", "addParagraphTemplate" , value, par.style)
	x
})



