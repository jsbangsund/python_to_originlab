# python_to_originlab
Functions in python to automate sending data and basic plotting functions in OriginLab

Origin color increment lists (.oth files) follow the color series in https://jiffyclub.github.io/palettable/, and can be found in the folder "\OriginTemplates\Themes\Graph".

There are a number of small syntax differences between versions. I have tested this code on OriginLab 2016 and 2018, but don't know about compatibility outside of those versions.

OriginLab's default directory for user templates is (with appropriate version year selected):
C:\Users\username\Documents\OriginLab\2016\User Files

The function matplotlib_to_origin will try it's best to convert a matplotlib figure to an origin graph, extracting data and line properties from the figure and axis handles. It's by no means perfect, but allows for data to quickly be transferred from the python environment to OriginLab for further tweaking and sharing.
![Example Origin Plot](/example.PNG)
