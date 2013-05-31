# tools functions 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 31 mars 2013
# Version: 0.1
###############################################################################
.onLoad= function(libname, pkgname){
#	if( is.null( options( "java.parameters" )$java.parameters ) )
#		options( java.parameters = c("-Xmx512M", "-Xms128M", "-Xss512K") )
	.jpackage( pkgname, lib.loc = libname )
	invisible()
}


.jinitialize.TableFormat = function( x ){
		obj = .jnew("com/lysis/docx4r/format/TableFormat", "%", 2L, 2L, 2L, "YYYY-mm-dd", "HH:MM", "YYYY-mm-dd HH:MM" )
		
		rootnames = c("header", "groupedheader", "double", "integer", "percent", "character", "date", "datetime")
		for( what in rootnames ){
			jwhatmethod = paste( casefold( substring(what , 1, 1 ),  upper = T ), casefold( substring(what , 2, nchar(what) ),  upper = F ), sep = "" )
			jwhatmethod = paste( "set", jwhatmethod, "Text", sep = "" )
			rwhatobject = paste( what, ".text", sep = "" )
			str = paste( "x@", rwhatobject , sep = "" )
			robject = eval( parse ( text = str ) )
			.jcall( obj, "V", jwhatmethod, as.character(robject@properties["color"])
					, as.integer(gsub("px$", "", robject@properties["font-size"]))
					, as.logical(robject@properties["font-weight"]=="bold")
					, as.logical(robject@properties["font-style"]=="italic")
					, as.character(robject@properties["font-family"])
			)
		
		}
		
		for( what in rootnames ){
			jwhatmethod = paste( casefold( substring(what , 1, 1 ),  upper = T ), casefold( substring(what , 2, nchar(what) ),  upper = F ), sep = "" )
			jwhatmethod = paste( "set", jwhatmethod, "Par", sep = "" )
			rwhatobject = paste( what, ".par", sep = "" )
			str = paste( "x@", rwhatobject , sep = "" )
			robject = eval( parse ( text = str ) )
			.jcall( obj, "V", jwhatmethod, as.character(robject@properties["text-align"])
					, as.integer( gsub("px$", "", robject@properties["padding-bottom"]) )
					, as.integer( gsub("px$", "", robject@properties["padding-top"]) )
					, as.integer( gsub("px$", "", robject@properties["padding-left"]) )
					, as.integer( gsub("px$", "", robject@properties["padding-right"]) )
			)
			
		}
		
		for( what in rootnames ){
			jwhatmethod = paste( casefold( substring(what , 1, 1 ),  upper = T ), casefold( substring(what , 2, nchar(what) ),  upper = F ), sep = "" )
			jwhatmethod = paste( "set", jwhatmethod, "Cell", sep = "" )
			rwhatobject = paste( what, ".cell", sep = "" )
			str = paste( "x@", rwhatobject , sep = "" )
			robject = eval( parse ( text = str ) )

			.jcall( obj, "V", jwhatmethod
					, as.character( robject@properties["border-bottom-color"] )
					, as.character( robject@properties["border-bottom-style"] )
					, as.integer( gsub("px$", "", robject@properties["border-bottom-width"] ) )
					, as.character( robject@properties["border-left-color"] )
					, as.character( robject@properties["border-left-style"] )
					, as.integer( gsub("px$", "", robject@properties["border-left-width"] ) )
					, as.character( robject@properties["border-top-color"] )
					, as.character( robject@properties["border-top-style"] )
					, as.integer( gsub("px$", "", robject@properties["border-top-width"] ) )
					, as.character( robject@properties["border-right-color"] )
					, as.character( robject@properties["border-right-style"] )
					, as.integer( gsub("px$", "", robject@properties["border-right-width"] ) )
					, as.character( robject@properties["vertical-align"] )
					, as.integer( gsub("px$", "", robject@properties["padding-bottom"] ) )
					, as.integer( gsub("px$", "", robject@properties["padding-top"] ) )
					, as.integer( gsub("px$", "", robject@properties["padding-left"] ) )
					, as.integer( gsub("px$", "", robject@properties["padding-right"] ) )
			)
			
		}
		obj
}

