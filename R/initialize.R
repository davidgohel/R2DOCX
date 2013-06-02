# Method Docx.initialize 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("initialize", "Docx", function(.Object, title, basefile) {
	if( missing( basefile ) )
		basefile = paste( find.package("R2DOCX"), "templates/EMPTY_DOC.docx", sep = "/" )
	.reg = regexpr( paste( "(\\.(?i)(docx))$", sep = "" ), basefile )
	
	if( !file.exists( basefile ) || .reg < 1 )
		stop(basefile , " is not a valid file.")

	if( missing( title ) ) title = ""
	
	obj = .jnew("com/lysis/reporting/docx4R" )
	.jcall( obj, "V", "setBaseDocument", basefile )
	.sysenv = Sys.getenv(c("USERDOMAIN","COMPUTERNAME","USERNAME"))
	
	.jcall( obj, "V", "setDocPropertyTitle", title )
	.jcall( obj, "V", "setDocPropertyCreator", paste( .sysenv["USERDOMAIN"], "/", .sysenv["USERNAME"], " on computer ", .sysenv["COMPUTERNAME"], sep = "" ) )
	
	.Object@obj = obj
	.Object@title = title
	.Object@basefile = basefile
	.Object@styles = .jcall( obj, "[S", "getStyleNames" )

	matchheaders = regexpr("^(Heading|Titre|Rubrik|Overskrift|berschrift)[1-9]{1}$", .Object@styles )
	if( any( matchheaders > 0 ) ){
		.Object <- setHeaderStyle(.Object
		, stylenames = sort( .Object@styles[ matchheaders > 0 ] ) )
	} else .Object@header.styles = character(0)
	
	.Object
})

