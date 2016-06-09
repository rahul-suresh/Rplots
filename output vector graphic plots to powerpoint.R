# code for generating vector grahic plot fo

# of course you can always create a tiff, jpeg or pdf and then place it manually in the graph.
# this way however you can use transparency in powerpoint if you need it. There are two ways that I can make this happen.


##output plot directly as a powerpoint slide using reporters
library(ReporteRs)
# create a new variable and defines it as supposed to be a pptx object. if run each time, it creates a new pptx file.
# remove this definition to just add new slide after defining it first time.
# if trying to add slide to existing presentation use this example code
## mydoc= pptx(template="D:/path to presentation")
mydoc = pptx() 

# There are a number of slide layouts to choose from this is the simplest slide layout for a data slide
mydoc = addSlide( mydoc, slide.layout = "Title and Content" ) 

# set the title of the graph
mydoc = addTitle( mydoc, "Plot examples" )  


mydoc = addPlot(mydoc, # variable to which the plot is added
	fun=function() print(Percentage), # function to add plot 
	vector.graphic=TRUE, # this will create a vector graphic plot in powerpoint vector graphics format that is in DrawingML format. if not specified it will create a png file.
	width = 5, height=5, # specify the width and the height of the plot. usually the slide is around 7.5 to 10inches in length so 5 inches each way is a good size for the plot in the slide.
	pointsize=40) # resets the text pointsize

writeDoc(mydoc, # specify variable
				 # create new file to adds all the attributes in the variable
				 file = "C:/Users/Rahul Suresh/OneDrive for Business/Data-Experiments/20160504 AFS98 test on mice/ggplot_cm.pptx"
				 )

##output plot directly as a powerpoint slide using dev emf

library(devEMF)

#create an emf file
emf(file="example.emf", # file name
		bg="white", # background
		width=12, height=8, # width and height of the file in inches
		family="Calibri", 
		pointsize=20)

plot(c(1:100), c(1:100), pch=20) #define the plot. in case of ggplot object, print it.

dev.off() # done!

# on-screen rendering of emf in powerpoint e.g. is not as good as done by reporters, as it renders non-anti-aliased
# transparency also isn't supported.
# Right clicking on the image and selecting Ungroup (or edit, i guess) converts it to native Office DrawingML format, which will then render anti-aliased
# but the placement of the text labels maybe slightly messed up then (can be solved by selecting and centering all text) 


## do not use win.metafile. it is awfull.
