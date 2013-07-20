# Method Docx.addParagraph 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("addParagraph", "Docx", function(x, value, stylename, replacements, replacement.styles, bookmark, ... ) {
	if( missing( stylename )) {
		stop("argument 'stylename' is missing")
	} else if( !is.element( stylename , styles( x ) ) ){
		stop(paste("Style {", stylename, "} does not exists.", sep = "") )
	}
	if( !missing( replacement.styles ) ){
		if( missing( replacements ) ) stop("argument 'replacements' is missing while argument 'replacement.styles' has been supplied")
		if( !all( sapply( replacement.styles, is, "textProperties" ) ) ) stop("argument 'replacement.styles' must be a list of 'textProperties' objects.")
	} else if( !missing( replacements ) ){
		stop("argument 'replacement.styles' is missing while argument 'replacement' has been supplied")
	}
	for( text.value in value ){
		paragrah = .jnew("com/lysis/docx4r/elements/Paragraph", text.value  )
		
		if( !missing( replacements ) ){

	
			# loop over each replacement
			for( j in names( replacements ) ){
				fs = replacement.styles[[j]]@properties["font-size"]
				fs = as.integer( gsub( "px", "", fs ) )
				.jcall( paragrah, "V", "addReplacement"
						, paste("[", j, "]", sep = "" )#keys
						, replacements[[j]] #value
						, fs #font-size
						, replacement.styles[[j]]@properties["font-weight"]=="bold"
						, replacement.styles[[j]]@properties["font-style"]=="italic"
						, replacement.styles[[j]]@properties["color"]
						, replacement.styles[[j]]@properties["font-family"]
				)
			}
		}
		if( missing( bookmark ) )
			.jcall( x@obj, "V", "add" , paragrah, stylename)
		else .jcall( x@obj, "V", "insert", bookmark, paragrah, stylename )

	}
	if( !missing( bookmark ) )
		.jcall( x@obj, "V", "deleteBookmark", bookmark )
	
	x
})



