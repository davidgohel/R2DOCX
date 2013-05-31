# Generic definitions
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

if(!isGeneric("styles"))
	setGeneric("styles", function(x, ...) standardGeneric("styles"));

if(!isGeneric("setHeaderStyle"))
	setGeneric("setHeaderStyle", function(x, stylenames) standardGeneric("setHeaderStyle"));

if(!isGeneric("replaceText"))
	setGeneric("replaceText", function(x, ...) standardGeneric("replaceText"));

if(!isGeneric("addTOC"))
	setGeneric("addTOC", function(x, ...) standardGeneric("addTOC"));

