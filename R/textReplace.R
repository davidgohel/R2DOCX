# Method Docx.textReplace 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("replaceText", "Docx", function(x, pattern , replacement, ...) {
			.jcall( x@obj, "V", "replaceText", pattern , replacement )
			x
		})


