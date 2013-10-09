# Method Docx.addImage 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 30 sept. 2013
# Version: 0.1
###############################################################################

setMethod("addImage", "Docx", function(x, filename, bookmark, ... ) {
			if( missing( filename ) )
				stop("filename cannot be missing")
			if( !file.exists( filename ) )
				stop( filename, " does not exist")
						
			# Send the graph to java that will 'encode64ize' and place it in a docx4J object
			if( missing( bookmark ) )
				.jcall( x@obj, "V", "addImage", .jarray( filename ) )
			else .jcall( x@obj, "V", "insertImage", bookmark, .jarray( filename ) )
			
			doc
		})


