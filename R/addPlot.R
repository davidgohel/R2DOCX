# Method Docx.addPlot 
# 
# Author: David GOHEL <david.gohel@lysis-consultants.fr>
# Date: 1 avr. 2013
# Version: 0.1
###############################################################################

setMethod("addPlot", "Docx", function(doc, fun
		, legend, stylename = "PlotReference"
		, width = 7, height = 6
		, pointsize = 12
		, visible = T, ... ) {
			if( !missing( legend) && legend != "" ){
				# check style does exist
				if( !is.element( stylename , styles( doc ) ) )
					stop(paste("Style {", stylename, "} does not exists. Use styles() to list available styles.", sep = "") )
			}
			plotargs = list(...)

			dirname = tempfile( )
			dir.create( dirname )
			filename = paste( dirname, "/plot%03d.png" ,sep = "" )

			grDevices::png (filename = filename
					, width = width*72.2, height = height*72.2
					, pointsize = pointsize
			)
			
			fun(...)
			dev.off()

			plotfiles = list.files( dirname , full.names = T )
			
			
			for( i in plotfiles )
			# Send the graph to java that will 'encode64' and place it in a docx4J object
				.jcall( doc@obj, "V", "addImage", i )

			# Finally, add the legend of the Graph after the Drawing
			if( !missing( legend) && legend != "" ) addParagraph( doc, value = legend, stylename = stylename )
			doc
		})

