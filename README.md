# NMDOT-Python-Tools
Python scripts and geoprocessing tools. 

I work in the GIS Unit at New Mexico Department of Transportation (NMDOT), supporting various DOT programs including drainage, environment, and planning. The code includes python toolboxes (ESRI ArcGIS vernacular) to aid the environment department when they do consultation with tribes, seeking their input on NMDOT projects that may have ground disturbance, as well as code we run from my group, GIS Support Unit, to automatically update a column in NMDOT's enterprise geodatabase when values are missing. 


1) Filter eSTIP Tables: This python tool was designed to increase efficiency for reviewing which NMDOT projects involve ground disturbance and may need tribal consultation prior to starting. It leverages pandas to combine multiple excel files that come from a NMDOT application using a common key. Once all necessary data is attained, it filters out project types with zero ground disturbance and outputs a formatted excel sheet with two sheets: one with all projects, and a second with the filtered reduced set that needs tribal consultation.

2) Tribal Consultation Area of Interest Maps & Tables: This python tool takes the filtered excel sheet and creates maps and tables depicting NMDOT projects in the counties of interest to those tribes.

3) Tribal Consultation Maps & Tables by County: This python tool takes the filtered excel sheet and creates maps and tables depicting NMDOT projects for individual counties and labels each project with its control number, so that tribes can easily find where specific projects are located. 

4) Route Converter: This python tool uses regular expressions to parse various components of the RouteID column found in most feature classes in NMDOT's enterprise geodatabase, and updates the DisplayID column if records are blank. 
