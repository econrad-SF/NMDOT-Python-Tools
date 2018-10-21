import arcpy, os, re, sys
import pandas as pd
#import numpy as np
#import matplotlib.pyplot as plt

# Global variables
nmdot_tools_version = 0.2
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
testing_eSTIP_tables = os.path.join(os.path.dirname(__file__), "Reference Files - Please don't alter\Environment\Tribal Consultation\Testing")
today = pd.datetime.now()

""" The Toolbox class contains all the various python tools that have been created. List of tools can be found in the self.tools propery. Each tool is it's own class (note how each tool's class name matches that
found in the Toolbox self.tools property. """
class Toolbox(object):
    def __init__(self):
        self.label = 'NMDOT GIS Tools'
        self.description = 'This toolbox contains custom GIS tools created specifically for NMDOT.'
        self.tools = [Filter_eSTIP_Tables, Tribal_AOI_Maps_Tables, Tribal_Consultation_Maps_Tables_by_County, RouteConverter]

""" Each tool defined in its own class """
class Filter_eSTIP_Tables(object):
    def __init__(self):
        self.label = '(1) Filter eSTIP Tables'
        self.description = "This tool takes the 3 tables: eSTIP Project Information export, eSTIP Tip Listing export, and the eSTIP Public portal export. It deletes extraneous columns, renames them,\
            replaces some record values (e.g. NM Dot becomes NMDOT), and removes extra STIP versions for the same control number by only keeping the most recent version. It then\
            merges the 3 eSTIP tables together using inner joins on the 'CN' field. Next it compares the full eSTIP table to Genevieve's inventory of projects that have previously been\
            consulted on. It does this by making a left outer join on the 'CN' field. Projects with a CN that matches one in the inventory remain in the 'eSTIP All' sheet but are dropped\
            in the 'eSTIP Filtered' sheet. At this point, the 'eSTIP All' sheet in the Excel output is complete. Next, a filter is applied to remove several project types that never cause\
            ground disturbance ('Administration (27)', 'Debt Service (45)', 'Planning (18)', 'Research (19)', 'ROW Acquisition (16)', 'Study/Planning (18)', and 'Training (42)'). This\
            completes the filtering process for the 'eSTIP Filtered' sheet.\
            Tool Output:\
            This tool returns an Excel file called Tribal_Consultation_[today's date].xlsx with two sheets:\
            1. eSTIP All       (Full output of the 3 merged Excel sheets to provide all the requried information of Geneveive Head in Environment)\
            2. eSTIP Filtered  (It excludes project type categories known to have zero ground disturbance and projects already consulted on.) "

        self.canRunInBackground = False
        self.category = 'Environment - Tribal Consultation'

    def getParameterInfo(self):
        """Define parameter definitions"""
        param0 = arcpy.Parameter(
            displayName = 'eSTIP Data Export -> Project Info (Excel File)',
            name = 'eSTIP_project_info',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param1 = arcpy.Parameter(
            displayName = 'eSTIP Reports -> Tip Listing (Excel File)',
            name = 'eSTIP_tip_listing',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param2 = arcpy.Parameter(
            displayName = 'Public eSTIP Portal Export (Excel File)',
            name = 'eSTIP_public_portal',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param3 = arcpy.Parameter(
            displayName = 'STIP Already Consulted (Excel File)',
            name = 'STIP_already_consulted',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param4 = arcpy.Parameter(
            displayName = 'Choose Output Location',
            name = 'output_location',
            datatype = 'DEFolder',
            parameterType = 'Required',
            direction = 'Input')
        param4.value = desktop
        param4.filter.list = ['File System']                                    # Set filter to only accept a folder
        param5 = arcpy.Parameter(
            displayName = 'Debugging Session?',
            name = 'debugging_session',
            datatype = 'GPBoolean',
            parameterType = 'Optional',
            direction = 'Input')
        params = [param0, param1, param2, param3, param4, param5]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        if parameters[5].value == True:
            parameters[0].value = os.path.join(testing_eSTIP_tables, 'eSTIP Data Export - Project Info.xlsx')
            parameters[1].value = os.path.join(testing_eSTIP_tables, 'eSTIP Reports - Tip Listing.xlsx')
            parameters[2].value = os.path.join(testing_eSTIP_tables, 'eSTIP - Public Export.xlsx')
            parameters[3].value = os.path.join(testing_eSTIP_tables, 'Already Consulted List.xlsx')
        elif parameters[5].value == False:
            parameters[0].value = ''
            parameters[1].value = ''
            parameters[2].value = ''
            parameters[3].value = ''

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

        # Ensure that user has correct version of Pandas installed. If not, cause error message to appear to prevent from running.
        if pd.__version__ not in ['0.23.0', '0.23.1', '0.23.2']:
            parameters[0].setErrorMessage(str('This tool requires Pandas 0.23.0, as well as xlsxwriter 1.0.5. Your ArcGIS python installation needs to be updated to include these Python packages. Please contact the NMDOT GIS Support Unit to perform the installation.'))

        return

    def execute(self, parameters, messages):

        # User-provdied input
        estip_project_info_table = parameters[0].valueAsText
        estip_tip_listing_table = parameters[1].valueAsText
        estip_public_table = parameters[2].valueAsText
        STIP_prev_consulted_table = parameters[3].valueAsText
        output_folder = parameters[4].valueAsText

        # 1st Table. Read in eSTIP Project Information .xlsx file (need to omit 1st row and very last row)
        estip = pd.read_excel(estip_project_info_table, skiprows = 1, skip_footer = 1, usecols = 56, dtype = {'VERSION': int, 'OLD ACCESS ID': str})  ########## # PANDAS CODE CHANGE from VERSION 0.16.1: estip = pd.read_excel(estip_table, skiprows = 1, skip_footer = 1) # ##########

        # 1st Table. Delete extraneous columns
        delete_cols = ['PROJECT ID', 'P CEMS NUM', 'EA NUM', 'STATE PJ ID', 'LOCAL PROJECT ID', 'FEDERAL PROJECT NUM', 'MTIP VERSION', 'MTIP VERSION2', 'ADMIN FORMAL', 'DOC YEAR',
            'MPO', 'MPO APPROVED', 'STATE APPROVE', 'FEDERAL APPROVE', 'STATE APPROVED', 'FEDERAL APPROVED', 'STATE STAFF APPROVED', 'FTA APPROVED', 'CAPACITY', 'TCM', 'AQ EXEMPT',
            'CMP', 'PRIMARY CONTACT', 'PHONE', 'EMAIL ADDRESS', 'MAP21 GOALS MEASURES', 'SYSTEM', 'LOCATION TYPE', 'local street name/', 'from/prim/near/xtreet/intchan',
            'to/second/xstreet/interchange', 'DIST MILES', 'BRIDGE NUMBER', 'NUM LOCS', 'CHANGE REASON', 'NARRATIVE', 'DATE SUBMITTED', 'SUBMITTED BY', 'APPROVED DATE', 'APPROVED BY',
            'ACCEPT COMMENTS', 'PROJECT CHANGES', 'PROGRAM AMOUNT', 'PROGRAM YEARS', 'FUNDING SOURCE']
        estip.drop(columns = delete_cols, inplace = True)

        # 1st Table. Rename column names
        estip.rename(columns = {'OLD ACCESS ID': 'CN', 'VERSION': 'Version', 'DISTRICT': 'District', 'COUNTY': 'County', 'LEAD AGENCY': 'Lead Agency', 'PROJECT TITLE': 'Project Title',
                    'PROJECT DESCRIPTION': 'Project Description', 'PRIMARY PROJECT TYPE': 'Project Type', 'ROUTE': 'Route', 'MILEPOST BEGIN': 'BOP', 'MILEPOST END': 'EOP', 'MILEPOST LENGTH': 'Length (Miles)'}, inplace = True)
        # 1st Table. Replace values
        estip.replace(to_replace = {'Lead Agency': {'NM Dot': 'NMDOT'}, 'District': {'District 1': 'D1', 'District 2': 'D2', 'District 3': 'D3', 'District 4': 'D4', 'District 5': 'D5', 'District 6': 'D6', 'Statewide': 'STA'}}, inplace = True)

        # 1st Table. Remove earlier versions of Control Numbers (CN). For instance, if there's a CN A301740 with 4 versions, delete the first 3.
        # 1st Table. Accomplish by sorting Version in descending order and then only keeping the first (keep highest version #)
        estip.sort_values(['CN', 'Version'], ascending = [True, False], inplace = True)
        estip.drop_duplicates(subset=['CN'], keep = 'first', inplace = True)

        # 2nd Table. Read in eSTIP fund information
        estip_tip_listing = pd.read_excel(estip_tip_listing_table, dtype = {'TIP ID': str})
        estip_tip_listing.rename(columns = {'TIP ID': 'CN'}, inplace = True)
        delete_cols = ['LOCAL ID', 'MPO', 'STIP', 'PROJECT TITLE', 'LEAD AGENCY', 'TYPE', 'PROJECT DESCRIPTION', 'PROJECT LIMITS', 'FED FUND', 'FED', 'STATE', 'LOC', 'PE', 'ROW', 'OTHER', 'TOTAL YEAR']
        estip_tip_listing.drop(columns = delete_cols, inplace = True)

        # 2nd Table. Drop Year columns (e.g. 2018, 2019, etc.) using Regular Expression accounting for fact that years will change over time. Syntax: ^ = start of string; \d = one digit (RegEx = drop any column beginning with a digit.)
        estip_tip_listing = estip_tip_listing[estip_tip_listing.columns.drop(list(estip_tip_listing.filter(regex='^\d')))]

        # 3rd Table. Read in eSTIP public download. Default name = 'export.xls'
        if estip_public_table[-4:] == '.xls':
            estip_public = pd.read_table(estip_public_table, engine = 'python')
        elif estip_public_table[-5:] == '.xlsx':
            estip_public = pd.read_excel(estip_public_table, dtype = {'Control #': str})
        estip_public.rename(columns = {'Control #': 'CN', 'FUNDS': 'Funds', 'FED YR': 'FY', 'PRODUCTION DATE': 'Production Date'}, inplace = True)

        # 4th Table: Read in the NMDOT projects that have already been consulted
        STIP_prev_consulted = pd.read_excel(STIP_prev_consulted_table, usecols = 0, dtype = {'CN': str})  # zero-based indexing & very first column is 'CN'

        # 4th Table: Add new column to be able to drop all of the STIP that has previously been consulted
        STIP_prev_consulted['Consulted'] = 'Yes'

        ### Join operations ###
        # Two Inner Joins for joining the three different eSTIP tables together on 'CN' field
        temp_merge = pd.merge(estip, estip_tip_listing, on = 'CN', how = 'inner')
        temp2_merge = pd.merge(temp_merge, estip_public, on = 'CN', how = 'inner')

        # Left Outer Join 4th table with 3 other eSTIP tables (i.e. temp2_merge)
        full_merge = pd.merge(temp2_merge, STIP_prev_consulted, how = 'left', on = 'CN')

        # Specify column order and sort values.
        full_merge = full_merge[['CN', 'District', 'County', 'Lead Agency', 'Project Title', 'Project Description', 'Project Type', 'Route', 'BOP', 'EOP', 'Length (Miles)', 'Funds', 'FY', 'Production Date', 'CON', 'Version', 'Consulted']]
        full_merge.sort_values(by = ['Consulted', 'Project Type', 'CN'], ascending = True, inplace = True)

        # Populate the consulted column of the remaining records with 'No';  full_merge includes all projects that haven't been consulted on, but it omits duplicate versions for same control number, which were removed using drop_duplicates function in prior step.
        def consulted(x):
            if x == 'Yes':
                return x
            else:
                return 'No'

        full_merge['Consulted'] = full_merge['Consulted'].apply(consulted)

        # Begin querying dataset to only select projects with potential for ground disturbance
        # Omit Project Types that NEVER have ground disturbance according to Genevieve Head; '~' is used to negate the .isin function (we don't want those categories to be included)
        df_reduced = full_merge[~full_merge['Project Type'].isin(['Administration (27)', 'Debt Service (45)', 'Intelligent Transportation Systems (24)', 'PE Oncall (18)', 'Planning (18)', 'Rail - Highway Grade Separation (22)', 'Research (19)', 'ROW Acquisition (16)', 'Study/Planning (18)', 'Training (42)', 'Transit (23)'])]

        # Drop any projects that have already been consulted (identified by 'Yes' in consulted field)
        df_reduced = df_reduced[df_reduced.Consulted != 'Yes']

        ### Begin code requiring xlsxwriter python package, which isn't a default package included with ArcGIS 10.5.1 (needs to be installed using pip) ###
        # Write each dataframe to a different worksheet and format.
        writer = pd.ExcelWriter(r'{}\Tribal_Consultation_{:%B_%d_%Y}.xlsx'.format(output_folder, today), engine = 'xlsxwriter', options = {'strings_to_numbers': True})
        workbook = writer.book
        full_merge.to_excel(writer, sheet_name='eSTIP All', index=False, freeze_panes=(1,1)) #full_merge.to_excel(r'{}\eSTIP_ALL_{}.xlsx'.format(output_folder, today), index=False, freeze_panes=(1,1)) ########## # PANDAS CODE CHANGE from VERSION 0.16.1: full_merge.to_csv(r'{}\eSTIP_All_{}.csv'.format(output_folder, today), index = False, encoding = 'utf8', columns = columns) #columns = ['CN', 'District', 'County', 'Lead Agency', 'Project Title', 'Project Description', 'Project Type', 'Route', 'BOP', 'EOP', 'Length (Miles)', 'Funds', 'Fed Yr', 'Production Date', 'CON']   to_excel function depends on openpyxl module, which isn't installed by default in ArcMap 10.4.1 # ##########
        df_reduced.to_excel(writer, sheet_name='eSTIP Filtered', index=False, freeze_panes=(1,1)) #df_reduced.to_excel(r'{}\eSTIP_Filtered_{}.xlsx'.format(output_folder, today), index=False, freeze_panes=(1,1)) ########## # PANDAS CODE CHANGE from VERSION 0.16.1: df_reduced.to_csv(r'{}\eSTIP_Filtered_{}.csv'.format(output_folder, today), index = False, encoding = 'utf8', columns = columns) # ##########
        worksheet = writer.sheets['eSTIP All']
        worksheet2 = writer.sheets['eSTIP Filtered']
        worksheet.set_column('A:C', 10); worksheet2.set_column('A:C', 10)
        worksheet.set_column('D:F', 20); worksheet2.set_column('D:F', 20)
        worksheet.set_column('G:G', 33); worksheet2.set_column('G:G', 33)
        worksheet.set_column('H:M', 10); worksheet2.set_column('H:M', 10)
        worksheet.set_column('N:N', 15); worksheet2.set_column('N:N', 15)
        worksheet.set_column('O:Q', 10); worksheet2.set_column('O:Q', 10)

        format2 = workbook.add_format({'num_format': '$#,###.00'})
        worksheet2.set_column('O:O', 13.5, format2)

        # Write the column headers with the defined format.
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'align': 'center',
            'valign': 'top',
            'fg_color': '#C0C0C0',
            'border': 1})

        # Apply header format to all columns
        for col_num, value in enumerate(full_merge.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for col_num, value in enumerate(df_reduced.columns.values):
            worksheet2.write(0, col_num, value, header_format)

        writer.save()  ### End code requiring xlsxwriter ###

        # Autmoatically load the Excel output file
        os.startfile(r'{}\Tribal_Consultation_{:%B_%d_%Y}.xlsx'.format(output_folder, today))
        return

class Tribal_AOI_Maps_Tables(object):
    def __init__(self):
        self.label = '(2) Tribal AOI eSTIP Maps & Tables'
        self.description = "This tool creates tribal consultation area of interest (AOI) maps for all tribes which NMDOT Environment consults regarding projects with ground disturbance potentital. The only input for the tool is a filtered\
            eSTIP Excel file that must have a sheet name titled 'eSTIP Filtered' containing STIP projects to be mapped. The tool adds the STIP projects to NMDOT's LRS as lines or points depending on project length, zooms to each AOI, and\
            converts all labels to annotation so that annotation can be moved below the county mask prior to exporting separate PDF maps for all 33 tribes. The tool provides the user control of specifying output location for the maps,\
            control of changing the map notes placed at the bottom of the map, & season text that is appended to the map title. The tool also exports a separate Excel file for each tribe containing the projects (if any) that are within the tribe's AOI."
        self.canRunInBackground = False
        self.category = 'Environment - Tribal Consultation'

    def getParameterInfo(self):
        """ Define parameter definitions """
        param0 = arcpy.Parameter(
            displayName = 'eSTIP List for Tribal Consultation (Excel File)',
            name = 'eSTIP_tribal_consultation',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param1 = arcpy.Parameter(
            displayName = 'Choose Output Location',
            name = 'output_location',
            datatype = 'DEFolder',
            parameterType = 'Required',
            direction = 'Input')
        param1.value = desktop
        param1.filter.list = ['File System']                                    # Set filter to only accept a folder
        param2 = arcpy.Parameter(
            displayName = 'Map Note Text (placed at map bottom)',
            name = 'map_note',
            datatype = 'GPString',
            parameterType = 'Required',
            direction = 'Input')
        param2.value = 'Map by NMDOT Environmental Bureau (GN Head)'
        param3 = arcpy.Parameter(
            displayName = 'Season (appended to Map Title)',
            name = 'season',
            datatype = 'GPString',
            parameterType = 'Required',
            direction = 'Input')
        param3.filter.list = ['Spring/Summer', 'Fall/Winter']
        if today.month in (1,2,3,4,5,6):
            param3.value = 'Spring/Summer'
        else:
            param3.value = 'Fall/Winter'
        param4 = arcpy.Parameter(
            displayName = 'Debugging Session?',
            name = 'debugging_session',
            datatype = 'GPBoolean',
            parameterType = 'Optional',
            direction = 'Input')
        params = [param0, param1, param2, param3, param4]
        return params

    def isLicensed(self):
        """ Set whether tool is licensed to execute. """
        return True

    def updateParameters(self, parameters):
        """ Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        if parameters[4].value == True:
            parameters[0].value = os.path.join(testing_eSTIP_tables, 'Tribal Consultation Filtered.xlsx')
        elif parameters[4].value == False:
            parameters[0].value = ''

        return

    def updateMessages(self, parameters):
        """ Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation. """
        # Ensure that user has correct version of Pandas installed. If not, cause error message to appear to prevent from running.
        if pd.__version__ not in ['0.23.0', '0.23.1', '0.23.2']:
            parameters[0].setErrorMessage(str('This tool requires Pandas 0.23.0, as well as xlsxwriter 1.0.5. Your ArcGIS python installation needs to be updated to include these Python packages. Please contact the NMDOT GIS Support Unit to perform the installation.'))
        else:
            pass
        return

    def execute(self, parameters, messages):

        # User's input variables and references to files
        start = pd.datetime.now()
        eSTIP_table = parameters[0].valueAsText
        output_location = parameters[1].valueAsText
        new_map_notes = parameters[2].valueAsText
        season = parameters[3].valueAsText
        reference_files_tc = os.path.join(os.path.dirname(__file__), "Reference Files - Please don't alter\Environment\Tribal Consultation")
        reference_shps = os.path.join(os.path.dirname(__file__), "Reference Files - Please don't alter\GDB\NMDOT_Base.gdb")
        ARNOLD = os.path.join(reference_shps, 'ARNOLD_RH_2018')
        county = os.path.join(reference_shps, 'counties')
        tribal_land = os.path.join(reference_shps, 'Tribal_Land')
        state = os.path.join(reference_shps, 'state')

        """Tribe dictionary object - Key = Tribe / Values

        Dict Values:
        1. Counties AOI                 <---- Update this dictionary value if tribes want to be consulted with for NMDOT projects in additional counties.
        2. Map Orientation              <---- With the current counties of interest, the map is either best at Portrait or Landscape layout.
        3. Tribe has Shapefile?         <---- Can we add tribal polygon to the map? (value is either 'Yes Shapefile' or 'No Shapefile'. Some tribes NMDOT consults with reside outside of New Mexico.
        4. Annotation Reference Scale   <---- Cartography subtlety that makes sure all the annotation across maps shows up in the same font.
        """

        tribal_dict = {
            'Apache Tribe of Oklahoma': [['Chaves', 'Curry', 'Lea', 'Quay', 'Roosevelt', 'San Miguel', 'Union'], 'Portrait', 'No Shapefile', '2576256'],
            'Commanche Nation': [['Chaves', 'Colfax', 'Curry', 'De Baca', 'Dona Ana', 'Eddy', 'Guadalupe', 'Harding', 'Lea', 'Lincoln', 'Los Alamos', 'Mora', 'Otero', 'Quay', 'Rio Arriba', 'Roosevelt', 'San Miguel', 'Sandoval', 'Santa Fe', 'Sierra', 'Socorro', 'Taos', 'Torrance', 'Union', 'Valencia'], 'Portrait', 'No Shapefile', '2690001'],
            'Fort Sill Apache Tribe': [['Catron', 'Dona Ana', 'Grant', 'Hidalgo', 'Luna', 'Sierra', 'Socorro'], 'Portrait', 'No Shapefile', '1682464'],
            'Hopi Tribe': [['Bernalillo', 'Catron', 'Cibola', 'Colfax', 'Dona Ana', 'Grant', 'Hidalgo', 'Los Alamos', 'Luna', 'McKinley', 'Mora', 'Rio Arriba', 'San Juan', 'San Miguel', 'Sandoval', 'Santa Fe', 'Sierra', 'Socorro', 'Taos', 'Torrance', 'Valencia'], 'Portrait', 'No Shapefile', '2927826'],
            'Jicarilla Apache Nation': [['Colfax', 'Guadalupe', 'Harding', 'Mora', 'Quay', 'Rio Arriba', 'San Miguel', 'Sandoval', 'Santa Fe', 'Taos', 'Torrance', 'Union'], 'Landscape', 'Yes Shapefile', '1911813'],
            'Kiowa Tribe': [['Chaves', 'Colfax', 'Curry', 'De Baca', 'Dona Ana', 'Eddy', 'Guadalupe', 'Harding', 'Lea', 'Lincoln', 'Mora', 'Otero', 'Quay', 'Rio Arriba', 'Roosevelt', 'San Juan', 'San Miguel', 'Santa Fe', 'Sierra', 'Socorro', 'Taos', 'Torrance', 'Union'], 'Portrait', 'No Shapefile', '2721623'],
            'Mescalero Apache Tribe': [['Catron', 'Chaves', 'Cibola', 'Colfax', 'Curry', 'De Baca', 'Dona Ana', 'Eddy', 'Grant', 'Guadalupe', 'Harding', 'Hidalgo', 'Lea', 'Lincoln', 'Luna', 'Mora', 'Otero', 'Quay', 'Roosevelt', 'San Miguel', 'Sierra', 'Socorro', 'Torrance', 'Union'], 'Portrait', 'Yes Shapefile', '2901638'],
            'Navajo Nation': [['Bernalillo', 'Catron', 'Cibola', 'De Baca', 'Dona Ana', 'Grant', 'Guadalupe', 'Los Alamos', 'McKinley', 'Mora', 'Rio Arriba', 'San Juan', 'San Miguel', 'Sandoval', 'Santa Fe', 'Sierra', 'Socorro', 'Taos', 'Torrance', 'Valencia'], 'Portrait', 'Yes Shapefile', '2712235'],
            'Pawnee Nation': [['Colfax', 'Curry', 'Guadalupe', 'Harding', 'Los Alamos', 'Mora', 'Quay', 'Rio Arriba', 'San Miguel', 'Sandoval', 'Santa Fe', 'Taos', 'Union'], 'Landscape', 'No Shapefile', '1876688'],
            'Pueblo of Acoma': [['Catron', 'Cibola', 'Grant', 'McKinley', 'San Juan', 'Socorro', 'Valencia'], 'Portrait', 'Yes Shapefile', '2654770'],
            'Pueblo of Cochiti': [['Los Alamos', 'San Miguel', 'Sandoval', 'Santa Fe'], 'Landscape', 'Yes Shapefile', '1350691'],
            'Pueblo of Isleta': [['Bernalillo', 'Cibola', 'De Baca', 'Dona Ana', 'Guadalupe', 'Lincoln', 'Otero', 'Sandoval', 'Sierra', 'Socorro', 'Torrance', 'Valencia'], 'Portrait', 'Yes Shapefile', '2328440'],
            'Pueblo of Jemez': [['Los Alamos', 'San Miguel', 'Sandoval'], 'Landscape', 'Yes Shapefile', '1350691'],
            'Pueblo of Laguna': [['Bernalillo', 'Catron', 'Cibola', 'McKinley', 'San Juan', 'Sandoval', 'Valencia'], 'Portrait', 'Yes Shapefile', '1969614'],
            'Pueblo of Nambe': [['Santa Fe'], 'Portrait', 'Yes Shapefile', '495380'],
            'Pueblo of Ohkay Owingeh': [['Bernalillo', 'Los Alamos', 'Rio Arriba', 'San Juan', 'Sandoval', 'Santa Fe', 'Taos'], 'Landscape', 'Yes Shapefile', '1516631'],
            'Pueblo of Picuris': [['Rio Arriba', 'Taos'], 'Landscape', 'Yes Shapefile', '816432'],
            'Pueblo of Pojoaque': [['Rio Arriba', 'Santa Fe'], 'Portrait', 'Yes Shapefile', '1018935'],
            'Pueblo of San Felipe': [['Bernalillo', 'Los Alamos', 'Sandoval', 'Santa Fe'], 'Portrait', 'Yes Shapefile', '855649'],
            'Pueblo of San Ildefonso': [['Los Alamos', 'Rio Arriba', 'Sandoval', 'Santa Fe'], 'Portrait', 'Yes Shapefile', '1018935'],
            'Pueblo of Sandia': [['Bernalillo', 'Sandoval'], 'Portrait', 'Yes Shapefile', '706298'],
            'Pueblo of Santa Ana': [['San Juan', 'Sandoval', 'Santa Fe'], 'Portrait', 'Yes Shapefile', '1491828'],
            'Pueblo of Santa Clara': [['Los Alamos', 'Rio Arriba', 'Sandoval', 'Santa Fe'], 'Portrait', 'Yes Shapefile', '1018935'],
            'Pueblo of Santo Domingo': [['Sandoval', 'Santa Fe'], 'Landscape', 'Yes Shapefile', '855649'],
            'Pueblo of Taos': [['Colfax', 'Mora', 'Rio Arriba', 'Taos'], 'Landscape', 'Yes Shapefile', '1216687'],
            'Pueblo of Tesuque': [['Chaves', 'Dona Ana', 'Eddy', 'Los Alamos', 'McKinley', 'Mora', 'Rio Arriba', 'San Juan', 'San Miguel', 'Sandoval', 'Santa Fe', 'Taos', 'Torrance'], 'Portrait', 'Yes Shapefile', '2712235'],
            'Pueblo of Ysleta del Sur': [['Bernalillo', 'Chaves', 'Dona Ana', 'Eddy', 'Lea', 'Lincoln', 'Luna', 'Otero', 'Sierra'], 'Landscape', 'No Shapefile', '2424171'],
            'Pueblo of Zia': [['Rio Arriba', 'San Juan', 'Sandoval'], 'Landscape', 'Yes Shapefile', '1285749'],
            'Pueblo of Zuni': [['Cibola', 'McKinley', 'San Miguel'], 'Landscape', 'Yes Shapefile', '1854139'],
            'Southern Ute Tribe': [['San Juan'], 'Landscape', 'No Shapefile', '718618'],
            'Ute Mountain Ute Tribe': [['San Juan'], 'Landscape', 'Yes Shapefile', '718618'],
            'White Mountain Apache Tribe': [['Bernalillo', 'Catron', 'Cibola', 'Dona Ana', 'Grant', 'Hidalgo', 'Luna', 'McKinley', 'Sierra', 'Socorro', 'Valencia'], 'Portrait', 'No Shapefile', '2415025'],
            'Wichita and Affiliated Tribes': [['San Miguel'], 'Landscape', 'No Shapefile', '708656']
            }

        # Create output folder for output location & create PDF & MXD folders
        os.mkdir(os.path.join(output_location, 'Tribal AOI Maps'))
        output_folder = os.path.join(output_location, 'Tribal AOI Maps')
        os.mkdir(os.path.join(output_folder, 'PDFs'))
        os.mkdir(os.path.join(output_folder, 'MXDs'))
        os.mkdir(os.path.join(output_folder, 'Tables'))

        # Modify route column from eSTIP_table so it matches route column in LRS
        df = pd.read_excel(eSTIP_table, sheet_name = 'eSTIP Filtered')
        df.Route = df.Route.str.replace(' ', '-')
        df.Route = df.Route + '-P'
        df.rename(columns = {'Length (Miles)': 'Length'}, inplace = True)

        # Drop records without route information (can't be mapped) & output formatted table in users chosen table
        df = df[df.Route.notnull()]
        df.to_excel(os.path.join(output_folder, 'eSTIP Filtered Table.xlsx'), index = False)

        # Create File GDB in users' chosen output location; import 'Tribal AOI Maps.xlsx'; create "Table View" as input for ArcGIS Make Route Layer function
        arcpy.SetProgressorLabel("Creating file geodatabase called 'Tribal_Consultation.gdb' at your chosen location: {}".format(output_folder))
        arcpy.CreateFileGDB_management(output_folder, out_name = 'Tribal_Consultation.gdb', out_version = 'CURRENT')
        arcpy.TableToTable_conversion(r'{}\eSTIP Filtered Table.xlsx\Sheet1$'.format(output_folder), out_path = r'{}\Tribal_Consultation.gdb'.format(output_folder), out_name = 'eSTIP_Filtered_table')
        arcpy.MakeTableView_management(r'{}\Tribal_Consultation.gdb\eSTIP_Filtered_table'.format(output_folder), 'events_lyr')

        # Add eSTIP to ARNOLD LRS (mapping eSTIP comes later however)
        # (1) Convert ARNOLD to feature layer
        arcpy.SetProgressorLabel('Adding eSTIP to NMDOT Linear Referencing System.')
        arcpy.MakeFeatureLayer_management(ARNOLD, 'LRS_lyr')

        # (2) Create route LINE event layer; Secondly, select & map only projects > 1000 feet (0.189394 mile) as lines
        arcpy.MakeRouteEventLayer_lr(in_routes = 'LRS_lyr', route_id_field = 'DisplayID', in_table = 'events_lyr', in_event_properties = 'Route LINE BOP EOP', out_layer = 'route_events_lyr', add_error_field = 'ERROR_FIELD')
        arcpy.MakeFeatureLayer_management('route_events_lyr', 'Long_Projects_temp', where_clause = """ "Length" > 0.189394 """)

        # (3) Create route POINT event layer (even line events are made into points); Secondly, select & map only projects shorter than 1000 feet (0.189394 mile) as points.
        arcpy.MakeRouteEventLayer_lr(in_routes = 'LRS_lyr', route_id_field = 'DisplayID', in_table = 'events_lyr', in_event_properties = 'Route POINT BOP', out_layer = 'route_events_as_points_lyr', add_error_field = 'ERROR_FIELD')
        arcpy.MakeFeatureLayer_management('route_events_as_points_lyr', 'Short_Projects_temp', where_clause = """ "Length" < 0.189394 """)

        # Save both event layers as feature classes, make them into permanent Feature Layers (source will be to Tribal_Consultation.gdb; temp Project layers above are temporary)
        arcpy.FeatureClassToFeatureClass_conversion('Long_Projects_temp', '{}\Tribal_Consultation.gdb'.format(output_folder), 'Long_Projects')
        arcpy.FeatureClassToFeatureClass_conversion('Short_Projects_temp', '{}\Tribal_Consultation.gdb'.format(output_folder), 'Short_Projects')

        # Make the permanent feature classes into feature layers for python (arcpy package) manipulation
        arcpy.MakeFeatureLayer_management(os.path.join(output_folder, r'Tribal_Consultation.gdb\Long_Projects'), 'Long_Projects')
        arcpy.MakeFeatureLayer_management(os.path.join(output_folder, r'Tribal_Consultation.gdb\Short_Projects'), 'Short_Projects')

        # Write header to Summary text file
        esri_dict = arcpy.GetInstallInfo()
        with open(os.path.join(output_folder, 'Summary_log.txt'), 'w') as f:
            f.write('Date: {: %B %d, %Y}\n'.format(today))
            f.write('Tool: Tribal AOI Maps\n')
            f.write('NMDOT GIS Tools Version: {}\n'.format(str(nmdot_tools_version)))
            f.write('ArcGIS Version: {}\n'.format(str(esri_dict['Version'])))
            f.write('Python Version: {}.{}.{}\n'.format(sys.version_info[0], sys.version_info[1], sys.version_info[2]))
            f.write('Pandas Version: {}\n\n\n'.format(pd.__version__))
            f.write('User Specified Input:\n')
            f.write('Map Notes: {}\n'.format(new_map_notes))
            f.write('Season Value: {}\n\n\n'.format(season))
            f.write('Tribal NMDOT/FHWA Project Summary:\n')

            # Loop through the tribes and make a unique map for each
            count = 1   # Counter updates in message to user as tool runs
            tribe_total = len(tribal_dict.keys())
            for key, values in sorted(tribal_dict.items()):

                # Extract key, values from dictionary
                tribe = key
                counties_list = values[0]
                map_orientation = values[1]
                tribes_w_shapefiles = values[2]
                annotation_ref_scale = values[3]

                # Mapping variables
                if map_orientation == 'Landscape':
                    mapdoc = arcpy.mapping.MapDocument(os.path.join(reference_files_tc, r'MXDs\tribal_consultation_landscape.mxd'))
                elif map_orientation == 'Portrait':
                    mapdoc = arcpy.mapping.MapDocument(os.path.join(reference_files_tc, r'MXDs\tribal_consultation_portrait.mxd'))

                main_df = arcpy.mapping.ListDataFrames(mapdoc)[0]
                inset_df = arcpy.mapping.ListDataFrames(mapdoc)[1]
                output_mxd = r'{}\MXDs\{} Map.mxd'.format(output_folder, tribe)

                # Create county ArcMap layer object based upon list of counties.
                # Create SQL statement first & then Make County Feature layer
                counties_list.insert(0, '')
                counties_list.append('')
                counties_sql = "', '".join(counties_list)
                counties_sql = counties_sql[3:-3]
                delimfield = arcpy.AddFieldDelimiters(county, "NAME")
                arcpy.MakeFeatureLayer_management(county, '{}_AOI_lyr'.format(tribe), where_clause = delimfield + " IN (" + counties_sql + ")")

                # Add tribal AOI (ie highlighted counties) to both main and inset data frames
                counties_AOI = arcpy.mapping.Layer('{}_AOI_lyr'.format(tribe))
                arcpy.mapping.AddLayer(main_df, counties_AOI, 'TOP')
                arcpy.mapping.AddLayer(inset_df, counties_AOI, 'BOTTOM')

                # Zoom Main Map extent to counties of interest
                map_extent = counties_AOI.getExtent(True)
                main_df.extent = map_extent
                main_df.scale = main_df.scale * 1.01   # Zoom out just a tad.

                # Update layer symbology in main & inset data frames
                county_main_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\counties_main.lyr'))
                county_main_update = arcpy.mapping.ListLayers(mapdoc, counties_AOI, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = county_main_update, source_layer = county_main_sym, symbology_only = True)
                county_inset_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\counties_inset.lyr'))
                county_inset_update = arcpy.mapping.ListLayers(mapdoc, counties_AOI, inset_df)[0]
                arcpy.mapping.UpdateLayer(inset_df, update_layer = county_inset_update, source_layer = county_inset_sym, symbology_only = True)

                # Create SQL statement
                counties_sql = "', '".join(counties_list)
                counties_sql = counties_sql[3:-3]

                # Turn off county mask layer for AOI (counties outside of AOI have white opaque color for clearly demarcating AOI)
                county_mask = arcpy.mapping.ListLayers(mapdoc, 'counties_mask', main_df)[0]
                county_mask.definitionQuery = ' "NAME" NOT IN (' + counties_sql + ')'

                # Spatial Analysis - determine # of projects in AOI (information will be placed in map text)
                arcpy.SelectLayerByLocation_management('Short_Projects', overlap_type = 'INTERSECT', select_features = '{}_AOI_lyr'.format(tribe))
                short_project_number = str(arcpy.GetCount_management('Short_Projects')[0])
                arcpy.SelectLayerByLocation_management('Long_Projects', overlap_type = 'INTERSECT', select_features = '{}_AOI_lyr'.format(tribe))
                long_project_number = str(arcpy.GetCount_management('Long_Projects')[0])

                # Begin code to create excel table summarizing projects in tribal AOI //////////////////////////////////////////////
                # Export selected projects in AOI as Excel sheets
                arcpy.TableToExcel_conversion(Input_Table = 'Short_Projects', Output_Excel_File = '{}\Tables\Short_Projects_AOI.xls'.format(output_folder))
                arcpy.TableToExcel_conversion(Input_Table = 'Long_Projects', Output_Excel_File = '{}\Tables\Long_Projects_AOI.xls'.format(output_folder))

                # Clear selections so nothing is highlighted when map is created
                arcpy.SelectLayerByAttribute_management('Short_Projects', 'CLEAR_SELECTION')
                arcpy.SelectLayerByAttribute_management('Long_Projects', 'CLEAR_SELECTION')

                # Read in tables, Add new column with project type (short or long), merge them, & export
                d1 = pd.read_excel('{}\Tables\Short_Projects_AOI.xls'.format(output_folder))
                d1['Map Symbol'] = 'Short Project'
                d2 = pd.read_excel('{}\Tables\Long_Projects_AOI.xls'.format(output_folder))
                d2['Map Symbol'] = 'Long Project'
                frames = [d1, d2]
                result = pd.concat(frames, sort = False)
                result.drop(columns = ['OBJECTID', 'LOC_ERROR', 'Shape_Length'], inplace = True)

                writer = pd.ExcelWriter(r'{}\Tables\{} eSTIP.xlsx'.format(output_folder, tribe))
                workbook = writer.book
                result.to_excel(writer, sheet_name = 'NMDOT Projects in AOI', index = False, freeze_panes = (1,1))
                worksheet = writer.sheets['NMDOT Projects in AOI']
                worksheet.set_column('A:C', 10)
                worksheet.set_column('D:F', 20)
                worksheet.set_column('G:G', 33)
                worksheet.set_column('H:M', 10)
                worksheet.set_column('N:N', 15)
                worksheet.set_column('O:Q', 10)
                worksheet.set_column('R:R', 13)

                # Write the column headers with the defined format.
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'align': 'center',
                    'valign': 'top',
                    'fg_color': '#C0C0C0',
                    'border': 1})

                # Apply header format to all columns
                for col_num, value in enumerate(result.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                writer.save()  # (End code for excel tabel summarization)    ///////////////////////////////////////////////////////////

                # Add Events to Map ///////
                # Events represented as lines (Projects longer than 1,000 ft)
                line_events_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\Long_Projects_AOI.lyr'))
                line_events = arcpy.mapping.Layer('Long_Projects')
                arcpy.mapping.AddLayer(main_df, line_events, 'TOP')
                line_update = arcpy.mapping.ListLayers(mapdoc, line_events, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = line_update, source_layer = line_events_sym, symbology_only = True)

                # Events represented as points (Projects less than 1,000 ft)
                point_events_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\Short_Projects_AOI.lyr'))
                point_events = arcpy.mapping.Layer('Short_Projects')
                arcpy.mapping.AddLayer(main_df, point_events, 'TOP')
                point_update = arcpy.mapping.ListLayers(mapdoc, point_events, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = point_update, source_layer = point_events_sym, symbology_only = True)

                # Modify Map Elements ///////
                legend_tribe = arcpy.mapping.ListLayoutElements(mapdoc, '', 'Tribe - Remove')[0]
                legend_county = arcpy.mapping.ListLayoutElements(mapdoc, '', 'County - Move')[0]
                if tribes_w_shapefiles == 'Yes Shapefile':
                    # Add tribe shapefile
                    delimfield = arcpy.AddFieldDelimiters(tribal_land, "name")
                    arcpy.MakeFeatureLayer_management(tribal_land, '{}_lyr'.format(tribe), where_clause = delimfield + " = '" + tribe + "'")
                    tribe_lyr = arcpy.mapping.Layer('{}_lyr'.format(tribe))
                    arcpy.mapping.AddLayer(main_df, tribe_lyr, 'TOP')
                    # Update tribe layer symbology in main data frame
                    tribal_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\tribal_land.lyr'))
                    tribe_update = arcpy.mapping.ListLayers(mapdoc, tribe_lyr, main_df)[0]
                    arcpy.mapping.UpdateLayer(main_df, update_layer = tribe_update, source_layer = tribal_sym, symbology_only = True)
                    tribe_update.transparency = 25; del tribal_sym

                    if str(mapdoc.title).strip() == 'Tribal Consultation Map - Portrait':
                        legend_tribe.elementPositionX = 4.8772
                        legend_tribe.elementPositionY = 0.5239
                        legend_county.elementPositionX = 4.8772
                        legend_county.elementPositionY = 0.2677
                    elif str(mapdoc.title).strip() == 'Tribal Consultation Map - Landscape':
                        legend_tribe.elementPositionX = 7.0699
                        legend_tribe.elementPositionY = 0.5509
                        legend_county.elementPositionX = 7.0699
                        legend_county.elementPositionY = 0.2845
                elif tribes_w_shapefiles == 'No Shapefile':
                    legend_tribe.elementPositionX = -3
                    if str(mapdoc.title).strip() == 'Tribal Consultation Map - Portrait':
                        legend_county.elementPositionX = 4.8772
                        legend_county.elementPositionY = 0.5239
                    elif str(mapdoc.title).strip() == 'Tribal Consultation Map - Landscape':
                        legend_county.elementPositionX = 7.0699
                        legend_county.elementPositionY = 0.5509

                new_title = 'NMDOT/FHWA Projects in Areas of Traditional Interest to the \n {}'.format(tribe)
                text_elements = arcpy.mapping.ListLayoutElements(mapdoc, 'TEXT_ELEMENT')
                for element in text_elements:
                    if element.name == 'Map Title':
                        element.text = new_title + ', ' + season + ' ' + str(today.year)
                    elif element.name == 'Tribal Legend':
                        element.text = tribe
                    elif element.name == 'Map Notes':
                        element.text = new_map_notes
                    elif element.name == 'Short Project Number Tribal AOI':
                        element.text = str(short_project_number)
                    elif element.name == 'Long Project Number Tribal AOI':
                        element.text = str(long_project_number)
                    elif element.name in ('County Legend', 'County Number of Projects'):
                        if tribe in ('Pueblo of Nambe', 'Southern Ute Tribe', 'Ute Mountain Ute Tribe'):
                            element.text = 'County of Interest'
                        else:
                            element.text = 'Counties of Interest'

                legend_county_projects = arcpy.mapping.ListLayoutElements(mapdoc, '', 'County Map Projects')[0]
                legend_tribal_projects = arcpy.mapping.ListLayoutElements(mapdoc, '', 'Tribal AOI Projects')[0]
                legend_nm_roads = arcpy.mapping.ListLayoutElements(mapdoc, '', 'NM Roads')[0]
                if str(mapdoc.title).strip() == 'Tribal Consultation Map - Portrait':
                    legend_county_projects.elementPositionX = -5
                    legend_tribal_projects.elementPositionX = 1.1874
                    legend_tribal_projects.elementPositionY = 0.2677
                    legend_nm_roads.elementPositionX = 4.8772
                elif str(mapdoc.title).strip() == 'Tribal Consultation Map - Landscape':
                    legend_tribal_projects.elementPositionX = 1.3518
                    legend_tribal_projects.elementPositionY = 0.2725
                    legend_county_projects.elementPositionX = -10
                    legend_nm_roads.elementPositionX = 5.5369

                # Save MXD to county specific one; This step is necessary prior to converting labels to annotation.
                mapdoc.saveACopy(output_mxd)

                # Convert labels to annotation //////////////////////////////////////////////////////////////
                # Use regular expression to replace two word county's space with underscore for annotation naming purposes
                pattern = re.compile('\s')
                tribe_underscore = pattern.sub('_', tribe)
                arcpy.SetProgressorLabel("Now converting labels to annotation for NMDOT Roads. Also turning off layer's dynamic labels.")

                # Unfortunately, this function converts ALL layers' labels to annotation even though we only want Route Annotation (will delete extraneous annotation layers in later step)
                arcpy.TiledLabelsToAnnotation_cartography(map_document = output_mxd, data_frame = 'Layers', polygon_index_layer = state, out_geodatabase = os.path.join(output_folder, 'Tribal_Consultation.gdb'), out_layer = 'Annotation_{}'.format(tribe), anno_suffix = '_{}_'.format(tribe_underscore), reference_scale_value = annotation_ref_scale, feature_linked = "STANDARD", generate_unplaced_annotation = "NOT_GENERATE_UNPLACED_ANNOTATION")

                # Create new map variables to reference county specific mxd
                mapdoc_final = arcpy.mapping.MapDocument(os.path.join(output_folder, r'MXDs\{} Map.mxd'.format(tribe)))
                df_final = arcpy.mapping.ListDataFrames(mapdoc_final)[0]

                # Add created annotation layer to map
                annotation_lyr = arcpy.mapping.Layer('Annotation_{}'.format(tribe))
                arcpy.mapping.AddLayer(df_final, annotation_lyr, 'Top')

                # Reorder layers so that the long & short projects are underneath county mask if they're outside of AOI
                for lyr in arcpy.mapping.ListLayers(mapdoc_final, '', df_final):
                    if lyr.name == '{}_lyr'.format(tribe):
                        move_Layer1 = lyr
                    elif lyr.name == '{}_AOI_lyr'.format(tribe):
                        move_Layer2 = lyr
                    elif lyr.name == 'Long_Projects':
                        move_Layer3 = lyr
                    elif lyr.name == 'Short_Projects':
                        move_Layer4 = lyr
                    elif lyr.name == 'Annotation_{}'.format(tribe):
                        move_Layer5 = lyr
                    elif lyr.name == 'counties_mask':
                        reference_layer = lyr
                        if arcpy.Exists('{}_lyr'.format(tribe)):      # <------- A check b/c lyr will only exist if tribe has shapefile
                            arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer1, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer2, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer3, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer4, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer5, 'AFTER')
                    elif lyr.name == 'NMDOT Roads':
                        lyr.showLabels = False

                    # Remove extraneous annotation layers, which were created by the arcpy.TiledLabelsToAnnotation_cartography() function
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\counties_labels_{}_{}'.format(tribe, tribe, annotation_ref_scale, tribe_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\USA_Southwest_labels_{}_{}'.format(tribe, tribe, annotation_ref_scale, tribe_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\Mexico_labels_{}_{}'.format(tribe, tribe, annotation_ref_scale, tribe_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)

                # Export mapdoc to PDF
                arcpy.SetProgressorLabel(r"Now saving '{} Tribal Consultation Map.pdf' at your chosen location: {}\PDFs".format(tribe, output_folder))
                arcpy.mapping.ExportToPDF(mapdoc_final, r'{}\PDFs\{} Tribal Consultation Map.pdf'.format(output_folder, tribe), image_quality = 'BEST')
                arcpy.AddMessage('{} NMDOT/FHWA Projects in AOI: (Short: {} Long: {})       [Tool Progress: {} of {} maps]'.format(tribe, short_project_number, long_project_number, str(count), str(tribe_total)))
                f.write('{} NMDOT/FHWA Projects in AOI: (Short: {} Long: {})       [Tool Progress: {} of {} maps]\n'.format(tribe, short_project_number, long_project_number, str(count), str(tribe_total)))

                # Save MXD & ensure main data frame is activated
                mapdoc_final.activeView = "PAGE_LAYOUT"
                arcpy.SetProgressorLabel(r"Now saving '{} Tribal Consultation Map.mxd' at your chosen location: {}\MXDs".format(tribe, output_folder))
                mapdoc_final.save()

                # Delete temp .xls files used to create summary '[Tribe] eSTIP.xls' files
                os.remove('{}\Tables\Short_Projects_AOI.xls'.format(output_folder))
                os.remove('{}\Tables\Long_Projects_AOI.xls'.format(output_folder))

                # Increase iterator
                count += 1

            end = pd.datetime.now()
            elapsed = end - start
            minutes, seconds = divmod(elapsed.total_seconds(), 60)
            f.write('\n\n')
            f.write('(Elapsed time: {:0>1} minutes {:.2f} seconds)'.format(int(minutes), seconds))

        # Delete modified Excel sheet that was used by each tribe in the for loop
        os.remove(os.path.join(output_folder, 'eSTIP Filtered Table.xlsx'))
        return

class Tribal_Consultation_Maps_Tables_by_County(object):
    def __init__(self):
        self.label = '(3) eSTIP Maps & Tables by County'
        self.description = "This tool creates a county focused map of eSTIP projects for any county with eSTIP projects (it excludes counties with 0 projects). The tool adds eSTIP projects to NMDOT's LRS\
            as either points or lines depending on project length (1000 ft is the demarcation). Projects are labeled by control number (CN attribue) and then all labels are converted to annotation so that\
            annotation can be moved below the county mask prior to export as a PDF into a folder created at run time. The tool provides the user control of specifying output location for the maps, control\
            of changing the map notes placed at the bottom of the map, & season text that is appended to the map title. The tool also exports a separate Excel file for each county containing projects occurring there."
        self.canRunInBackground = False
        self.category = 'Environment - Tribal Consultation'

    def getParameterInfo(self):
        """ Define parameter definitions """
        param0 = arcpy.Parameter(
            displayName = 'eSTIP List for Tribal Consultation (Excel File)',
            name = 'eSTIP_tribal_consultation',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param1 = arcpy.Parameter(
            displayName = 'Choose Output Location',
            name = 'output_location',
            datatype = 'DEFolder',
            parameterType = 'Required',
            direction = 'Input')
        param1.value = desktop
        param1.filter.list = ['File System']                                    # Set filter to only accept a folder
        param2 = arcpy.Parameter(
            displayName = 'Map Note Text (placed at map bottom)',
            name = 'map_note',
            datatype = 'GPString',
            parameterType = 'Required',
            direction = 'Input')
        param2.value = 'Map by NMDOT Environmental Bureau (GN Head)'
        param3 = arcpy.Parameter(
            displayName = 'Season (appended to Map Title)',
            name = 'season',
            datatype = 'GPString',
            parameterType = 'Required',
            direction = 'Input')
        param3.filter.list = ['Spring/Summer', 'Fall/Winter']
        if today.month in (1,2,3,4,5,6):
            param3.value = 'Spring/Summer'
        else:
            param3.value = 'Fall/Winter'
        param4 = arcpy.Parameter(
            displayName = 'Debugging Session?',
            name = 'debugging_session',
            datatype = 'GPBoolean',
            parameterType = 'Optional',
            direction = 'Input')
        params = [param0, param1, param2, param3, param4]
        return params

    def isLicensed(self):
        """ Set whether tool is licensed to execute. """
        return True

    def updateParameters(self, parameters):
        """ Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        if parameters[4].value == True:
            parameters[0].value = os.path.join(testing_eSTIP_tables, 'Tribal Consultation Filtered.xlsx')
        elif parameters[4].value == False:
            parameters[0].value = ''
        return

    def updateMessages(self, parameters):
        """ Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation. """

        # Ensure that user has correct version of Pandas installed. If not, cause error message to appear to prevent from running.
        if pd.__version__ not in ['0.23.0', '0.23.1']:
            parameters[0].setErrorMessage(str('This tool requires Pandas 0.23.0, as well as xlsxwriter 1.0.5. Your ArcGIS python installation needs to be updated to include these Python packages. Please contact the NMDOT GIS Support Unit to perform the installation.'))
        else:
            pass
        return

    def execute(self, parameters, messages):

        # User's input variables and references to files
        start = pd.datetime.now()
        eSTIP_table = parameters[0].valueAsText
        output_location = parameters[1].valueAsText
        new_map_notes = parameters[2].valueAsText
        season = parameters[3].valueAsText
        reference_files_tc = os.path.join(os.path.dirname(__file__), "Reference Files - Please don't alter\Environment\Tribal Consultation")
        reference_shps = os.path.join(os.path.dirname(__file__), "Reference Files - Please don't alter\GDB\NMDOT_Base.gdb")
        ARNOLD = os.path.join(reference_shps, 'ARNOLD_RH_2018')
        county = os.path.join(reference_shps, 'counties')
        state = os.path.join(reference_shps, 'state')

        """ County dictionary object - Key = County / Values

        Values:
        1. Map orientation               <---- Update these if necessary
        2. Annotation Reference Scale    <---- Update these if necessary (Reference scale used in arcpy.TiledLabelsToAnnotation_cartography function)
        """
        county_dict = {
            'Bernalillo': ['Landscape', '387254'],
            'Catron': ['Portrait', '788788'],
            'Chaves': ['Portrait', '928504'],
            'Cibola': ['Landscape', '744213'],
            'Colfax': ['Landscape', '589243'],
            'Curry': ['Portrait', '369846'],
            'De Baca': ['Landscape', '590640'],
            'Dona Ana': ['Portrait', '710081'],
            'Eddy': ['Landscape', '730003'],
            'Grant': ['Portrait', '766928'],
            'Guadalupe': ['Landscape', '656204'],
            'Harding': ['Landscape', '620893'],
            'Hidalgo': ['Portrait', '822752'],
            'Lea': ['Portrait', '884153'],
            'Lincoln': ['Landscape', '914783'],
            'Los Alamos': ['Portrait', '123048'],
            'Luna': ['Portrait', '482389'],
            'McKinley': ['Landscape', '799780'],
            'Mora': ['Landscape', '512657'],
            'Otero': ['Portrait', '782493'],
            'Quay': ['Portrait', '642433'],
            'Rio Arriba': ['Landscape', '820160'],
            'Roosevelt': ['Portrait', '578000'],
            'San Juan': ['Landscape', '782208'],
            'San Miguel': ['Landscape', '771804'],
            'Sandoval': ['Landscape', '782599'],
            'Santa Fe': ['Portrait', '539523'],
            'Sierra': ['Landscape', '668055'],
            'Socorro': ['Landscape', '846451'],
            'Taos': ['Portrait', '552019'],
            'Torrance': ['Landscape', '591594'],
            'Union': ['Portrait', '709199'],
            'Valencia': ['Landscape', '400417']
            }

        # Create output folder for output location & create PDF & MXD folders
        os.mkdir(os.path.join(output_location, 'Tribal Consultation Maps by County'))
        output_folder = os.path.join(output_location, 'Tribal Consultation Maps by County')
        os.mkdir(os.path.join(output_folder, 'PDFs'))
        os.mkdir(os.path.join(output_folder, 'MXDs'))
        os.mkdir(os.path.join(output_folder, 'Tables'))

        # Modify route column from eSTIP_table so it matches DisplayID column in LRS
        df = pd.read_excel(eSTIP_table, sheet_name = 'eSTIP Filtered')
        df.Route = df.Route.str.replace(' ', '-')
        df.Route = df.Route + '-P'
        df.rename(columns = {'Length (Miles)': 'Length'}, inplace = True)

        # Drop records without route information (b/c they can't be mapped)
        df = df[df.Route.notnull()]
        df.to_excel(os.path.join(output_folder, 'eSTIP Filtered Table.xlsx'), index = False)

        # Create File GDB in users' chosen output location; import 'Tribal AOI Maps.xlsx'; create "Table View" as input for ArcGIS Make Route Layer function
        arcpy.SetProgressorLabel("Creating file geodatabase called 'Tribal_Consultation.gdb' at your chosen location: {}".format(output_folder))
        arcpy.CreateFileGDB_management(output_folder, out_name = 'Tribal_Consultation.gdb', out_version = 'CURRENT')
        arcpy.TableToTable_conversion(r'{}\eSTIP Filtered Table.xlsx\Sheet1$'.format(output_folder), out_path = r'{}\Tribal_Consultation.gdb'.format(output_folder), out_name = 'eSTIP_Filtered_table')
        arcpy.MakeTableView_management(r'{}\Tribal_Consultation.gdb\eSTIP_Filtered_table'.format(output_folder), 'events_lyr')

        # Add eSTIP to ARNOLD LRS (mapping eSTIP comes later however)
        # (1) Convert ARNOLD to feature layer
        arcpy.SetProgressorLabel('Adding eSTIP to NMDOT Linear Referencing System.')
        arcpy.MakeFeatureLayer_management(ARNOLD, 'LRS_lyr')

        # (2) Create route LINE event layer; Secondly, select & map only projects > 1000 feet (0.189394 mile) as lines
        arcpy.MakeRouteEventLayer_lr(in_routes = 'LRS_lyr', route_id_field = 'DisplayID', in_table = 'events_lyr', in_event_properties = 'Route LINE BOP EOP', out_layer = 'route_events_lyr', add_error_field = 'ERROR_FIELD')
        arcpy.MakeFeatureLayer_management('route_events_lyr', 'Long_Projects_temp', where_clause = """ "Length" > 0.189394 """)

        # (3) Create route POINT event layer (even line events are made into points); Secondly, select & map only projects shorter than 1000 feet (0.189394 mile) as points.
        arcpy.MakeRouteEventLayer_lr(in_routes = 'LRS_lyr', route_id_field = 'DisplayID', in_table = 'events_lyr', in_event_properties = 'Route POINT BOP', out_layer = 'route_events_as_points_lyr', add_error_field = 'ERROR_FIELD')
        arcpy.MakeFeatureLayer_management('route_events_as_points_lyr', 'Short_Projects_temp', where_clause = """ "Length" < 0.189394 """)

        # Save both event layers as feature classes
        arcpy.FeatureClassToFeatureClass_conversion('Long_Projects_temp', '{}\Tribal_Consultation.gdb'.format(output_folder), 'Long_Projects')
        arcpy.FeatureClassToFeatureClass_conversion('Short_Projects_temp', '{}\Tribal_Consultation.gdb'.format(output_folder), 'Short_Projects')
        arcpy.MakeFeatureLayer_management(os.path.join(output_folder, r'Tribal_Consultation.gdb\Long_Projects'), 'Long_Projects')
        arcpy.MakeFeatureLayer_management(os.path.join(output_folder, r'Tribal_Consultation.gdb\Short_Projects'), 'Short_Projects')

        # Select the counties with eSTIP Projects (only maps with eSTIP Projects within them will be mapped)
        arcpy.MakeFeatureLayer_management(county, "county_lyr")
        arcpy.SelectLayerByLocation_management("county_lyr", overlap_type = "INTERSECT", select_features = "Long_Projects", selection_type = "NEW_SELECTION")
        arcpy.SelectLayerByLocation_management("county_lyr", overlap_type = "INTERSECT", select_features = "Short_Projects", selection_type = "ADD_TO_SELECTION")

        # Create cursor to retrieve name of the counties with eSTIP projects. Manipule cursor object to create list of counties that have eSTIP
        cursor = arcpy.da.SearchCursor("county_lyr", "NAME")                    # returns an iterator of tuples.
        counties_str = ''.join(map(str, cursor))                                # convert to string
        pattern = re.compile('(\(u)+')                                          # matches '(u' when they're grouped
        counties_str = pattern.sub("", counties_str)
        counties_str = counties_str.replace(')', '')
        counties_str = counties_str.replace("'", "")
        counties_list = counties_str.split(',')
        counties_list = counties_list[:-1]                                      # drop last item which is ','

        # Write header to Summary text file
        esri_dict = arcpy.GetInstallInfo()
        with open(os.path.join(output_folder, 'Summary_log.txt'), 'w') as f:
            f.write('Date: {: %B %d, %Y}\n'.format(today))
            f.write('Tool: Tribal Consultation Maps by County\n')
            f.write('NMDOT GIS Tools Version: {}\n'.format(str(nmdot_tools_version)))
            f.write('ArcGIS Version: {}\n'.format(str(esri_dict['Version'])))
            f.write('Python Version: {}.{}.{}\n'.format(sys.version_info[0], sys.version_info[1], sys.version_info[2]))
            f.write('Pandas Version: {}\n\n\n'.format(pd.__version__))
            f.write('User Specified Input:\n')
            f.write('Map Notes: {}\n'.format(new_map_notes))
            f.write('Season Value: {}\n\n\n'.format(season))
            f.write('County NMDOT/FHWA Project Summary:\n')

            # Loop through counties with eSTIP projects and create a PDF map for each
            count = 1
            county_total = len(counties_list)
            for c in sorted(counties_list):

                map_orientation = county_dict[c][0]
                annotation_ref_scale = county_dict[c][1]

                # Determine whether map document has Landscape or Portrait layout
                if map_orientation == 'Landscape':
                    mapdoc = arcpy.mapping.MapDocument(os.path.join(reference_files_tc, r'MXDs\tribal_consultation_landscape.mxd'))
                elif map_orientation == 'Portrait':
                    mapdoc = arcpy.mapping.MapDocument(os.path.join(reference_files_tc, r'MXDs\tribal_consultation_portrait.mxd'))

                # Map variables
                main_df = arcpy.mapping.ListDataFrames(mapdoc)[0]
                inset_df = arcpy.mapping.ListDataFrames(mapdoc)[1]
                output_mxd = r'{}\MXDs\{} County.mxd'.format(output_folder, c)
                county_main_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\counties_main.lyr'))
                county_inset_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\counties_inset.lyr'))
                county_main_df = arcpy.mapping.ListLayers(mapdoc, "counties_mask", main_df)[0]
                arcpy.SelectLayerByAttribute_management(county_main_df, selection_type = "NEW_SELECTION", where_clause = """ "NAME" = '""" + c + """'""")

                # Zoom Main Map extent to county of interest
                main_df.zoomToSelectedFeatures()

                # Add county of interest to main & inset data frame
                arcpy.MakeFeatureLayer_management(county_main_df, '{}_county'.format(c))
                county_aoi = arcpy.mapping.Layer('{}_county'.format(c))
                arcpy.mapping.AddLayer(main_df, county_aoi, 'TOP')
                arcpy.mapping.AddLayer(inset_df, county_aoi, 'BOTTOM')

                # Update county layer symbology in main & inset data frames
                county_main_update = arcpy.mapping.ListLayers(mapdoc, county_aoi, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = county_main_update, source_layer = county_main_sym, symbology_only = True)
                county_inset_update = arcpy.mapping.ListLayers(mapdoc, county_aoi, inset_df)[0]
                arcpy.mapping.UpdateLayer(inset_df, update_layer = county_inset_update, source_layer = county_inset_sym, symbology_only = True)

                # Spatial Analysis - determine # of project in counties that have projects (information will be placed in map)
                arcpy.SelectLayerByLocation_management('Short_Projects', overlap_type = 'INTERSECT', select_features = '{}_county'.format(c))
                short_project_number = str(arcpy.GetCount_management('Short_Projects')[0])
                arcpy.SelectLayerByLocation_management('Long_Projects', overlap_type = 'INTERSECT', select_features = '{}_county'.format(c))
                long_project_number = str(arcpy.GetCount_management('Long_Projects')[0])

                # Create Excel table of projects in Tribal AOI ///////////////////////////////////////////////////////////
                # Export selected projects in AOI as Excel sheets
                arcpy.TableToExcel_conversion(Input_Table = 'Short_Projects', Output_Excel_File = '{}\Tables\{}_Short_Projects.xls'.format(output_folder, c))
                arcpy.TableToExcel_conversion(Input_Table = 'Long_Projects', Output_Excel_File = '{}\Tables\{}_Long_Projects.xls'.format(output_folder, c))

                # Clear selections so nothing is highlighted when map is created
                arcpy.SelectLayerByAttribute_management(county_main_df, 'CLEAR_SELECTION')
                arcpy.SelectLayerByAttribute_management('Short_Projects', 'CLEAR_SELECTION')
                arcpy.SelectLayerByAttribute_management('Long_Projects', 'CLEAR_SELECTION')

                # Read in tables, Add new column with project type (short or long), merge them, & export
                d1 = pd.read_excel('{}\Tables\{}_Short_Projects.xls'.format(output_folder, c))
                d1['Map Symbol'] = 'Short Project'
                d2 = pd.read_excel('{}\Tables\{}_Long_Projects.xls'.format(output_folder, c))
                d2['Map Symbol'] = 'Long Project'
                frames = [d1, d2]
                result = pd.concat(frames, sort = False)
                result.drop(columns = ['OBJECTID', 'LOC_ERROR', 'Shape_Length'], inplace = True)

                writer = pd.ExcelWriter(r'{}\Tables\{} County eSTIP.xlsx'.format(output_folder, c))
                workbook = writer.book
                result.to_excel(writer, sheet_name = '{} County Projects'.format(c), index = False, freeze_panes = (1,1))
                worksheet = writer.sheets['{} County Projects'.format(c)]
                worksheet.set_column('A:C', 10)
                worksheet.set_column('D:F', 20)
                worksheet.set_column('G:G', 33)
                worksheet.set_column('H:M', 10)
                worksheet.set_column('N:N', 15)
                worksheet.set_column('O:Q', 10)
                worksheet.set_column('R:R', 13)

                # Write the column headers with the defined format.
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'align': 'center',
                    'valign': 'top',
                    'fg_color': '#C0C0C0',
                    'border': 1})

                # Apply header format to all columns
                for col_num, value in enumerate(result.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                writer.save()

                # Turn off county mask layer for AOI
                c_sql = "'" + c + "'"
                county_mask = arcpy.mapping.ListLayers(mapdoc, 'counties_mask', main_df)[0]
                county_mask.definitionQuery = ' "NAME" <> ' + c_sql

                # Modify Map Elements
                new_title = 'NMDOT/FHWA Projects in {} County'.format(c)
                text_elements = arcpy.mapping.ListLayoutElements(mapdoc, 'TEXT_ELEMENT')
                for element in text_elements:
                    if element.name == 'Map Title':
                        element.text = new_title + ', ' + '\n' + season + ' ' + str(today.year)
                    elif element.name == 'Map Notes':
                        element.text = new_map_notes
                    elif element.name == 'County Legend':
                        element.text = '{} County'.format(c)
                    elif element.name == 'Short Project Number County Maps':
                        element.text = str(short_project_number)
                    elif element.name == 'Long Project Number County Maps':
                        element.text = str(long_project_number)
                    elif element.name == 'County Number of Projects':
                        element.text = '# of Projects in\n{} County'.format(c)
                legend_county_projects = arcpy.mapping.ListLayoutElements(mapdoc, '', 'County Map Projects')[0]
                legend_aoi_projects = arcpy.mapping.ListLayoutElements(mapdoc, '', 'Tribal AOI Projects')[0]
                legend_nm_roads = arcpy.mapping.ListLayoutElements(mapdoc, '', 'NM Roads')[0]
                legend_tribe = arcpy.mapping.ListLayoutElements(mapdoc, '', 'Tribe - Remove')[0]
                legend_county = arcpy.mapping.ListLayoutElements(mapdoc, '', 'County - Move')[0]
                legend_tribe.elementPositionX = -3
                if str(mapdoc.title).strip() == 'Tribal Consultation Map - Portrait':
                    legend_aoi_projects.elementPositionX = -5
                    legend_county_projects.elementPositionX = 1.1521
                    legend_county_projects.elementPositionY = 0.2677
                    legend_nm_roads.elementPositionX = 5.3367
                    legend_county.elementPositionX = 5.3367
                    legend_county.elementPositionY = 0.5018
                elif str(mapdoc.title).strip() == 'Tribal Consultation Map - Landscape':
                    legend_aoi_projects.elementPositionX = -5
                    legend_county_projects.elementPositionX = 1.1465
                    legend_county_projects.elementPositionY = 0.2725
                    legend_county.elementPositionY = 0.5179
                    legend_nm_roads.elementPositionX = 5.6387

                # Add Events to Map
                # Events represented as lines
                line_events_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\Long_Projects_County.lyr'))    # <----- NOTE: If labels break due to an ESRI update, you may need to recreate the .lyr file with all the label settings!
                line_events = arcpy.mapping.Layer('Long_Projects')
                arcpy.mapping.AddLayer(main_df, line_events, 'TOP')
                line_update_layer = arcpy.mapping.ListLayers(mapdoc, line_events, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = line_update_layer, source_layer = line_events_sym, symbology_only = False)

                # Events represented as points
                point_events_sym = arcpy.mapping.Layer(os.path.join(reference_files_tc, r'Layers\Short_Projects_County.lyr'))     # <----- NOTE: If labels break due to an ESRI update, you may need to recreate the .lyr file with all the label settings!
                point_events = arcpy.mapping.Layer('Short_Projects')
                arcpy.mapping.AddLayer(main_df, point_events, 'TOP')
                point_update_layer = arcpy.mapping.ListLayers(mapdoc, point_events, main_df)[0]
                arcpy.mapping.UpdateLayer(main_df, update_layer = point_update_layer, source_layer = point_events_sym, symbology_only = False)

                # Save MXD to county specific one; This step is necessary prior to converting labels to annotation. Without it, saving labels to annotation wouldn't include STIP labels (which aren't included in the reference mxds to create these maps i.e. tribal_consultation_[landscape,portrait].mxd)
                mapdoc.saveACopy(output_mxd)

                # Convert labels to annotation
                # Use regular expression to replace two word county's space with underscore for annotation naming purposes
                pattern = re.compile('\s')
                county_underscore = pattern.sub('_', c)
                arcpy.SetProgressorLabel("Now converting labels to annotation for (1) Short Projects (2) Long Projects and (3) NMDOT Roads. Also turning off these layers' dynamic labels.")
                arcpy.TiledLabelsToAnnotation_cartography(map_document = output_mxd, data_frame = 'Layers', polygon_index_layer = state, out_geodatabase = os.path.join(output_folder, 'Tribal_Consultation.gdb'), out_layer = 'Annotation_{}'.format(c), anno_suffix = '_{}_'.format(county_underscore), reference_scale_value = annotation_ref_scale, feature_linked = 'STANDARD', generate_unplaced_annotation = 'NOT_GENERATE_UNPLACED_ANNOTATION')

                # Create new map variables to reference county specific mxd
                mapdoc_final = arcpy.mapping.MapDocument(os.path.join(output_folder, r'MXDs\{} County.mxd'.format(c)))
                df_final = arcpy.mapping.ListDataFrames(mapdoc_final)[0]

                # Add created annotation layer to map
                annotation_lyr = arcpy.mapping.Layer('Annotation_{}'.format(c))
                arcpy.mapping.AddLayer(df_final, annotation_lyr, 'Top')

                # Reorder/remove/turn off layer labels
                for lyr in arcpy.mapping.ListLayers(mapdoc_final, '', df_final):
                    if lyr.name == '{}_county'.format(c):
                        move_Layer1 = lyr
                    elif lyr.name == 'Long_Projects':
                        lyr.showLabels = False
                        move_Layer2 = lyr
                    elif lyr.name == 'Short_Projects':
                        lyr.showLabels = False
                        move_Layer3 = lyr
                    elif lyr.name == 'Annotation_{}'.format(c):
                        move_Layer4 = lyr
                    elif lyr.name == 'counties_mask':
                        reference_layer = lyr
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer1, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer2, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer3, 'AFTER')
                        arcpy.mapping.MoveLayer(df_final, reference_layer, move_Layer4, 'AFTER')
                    elif lyr.name == 'NMDOT Roads':
                        lyr.showLabels = False

                    # Remove extraneous annotation layers, which were created by the arcpy.TiledLabelsToAnnotation_cartography() function
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\counties_labels_{}_{}'.format(c, c, annotation_ref_scale, county_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\USA_Southwest_labels_{}_{}'.format(c, c, annotation_ref_scale, county_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)
                    elif lyr.longName == r'Annotation_{}\Annotation_{}{}\Mexico_labels_{}_{}'.format(c, c, annotation_ref_scale, county_underscore, annotation_ref_scale):
                        arcpy.mapping.RemoveLayer(df_final, lyr)

                # Export mapdoc_final to PDF
                mapdoc_final.activeView = 'PAGE_LAYOUT'
                mapdoc_final.save()
                arcpy.SetProgressorLabel(r"Now saving '{} County Tribal Consultation.pdf' at your chosen location: {}\PDFs".format(c, output_folder))
                arcpy.mapping.ExportToPDF(mapdoc_final, r'{}\PDFs\{} County Tribal Consultation.pdf'.format(output_folder, c), image_quality = 'BEST')
                arcpy.AddMessage('{} County NMDOT/FHWA Projects: (Short: {} Long: {})     [Tool Progress: {} of {} maps]'.format(c, short_project_number, long_project_number, str(count), str(county_total)))
                f.write('{} County NMDOT/FHWA Projects: (Short: {} Long: {})     [Tool Progress: {} of {} maps]\n'.format(c, short_project_number, long_project_number, str(count), str(county_total)))

                # Delete temp .xls files used to create summary '[Tribe] eSTIP.xls' files
                os.remove('{}\Tables\{}_Short_Projects.xls'.format(output_folder, c))
                os.remove('{}\Tables\{}_Long_Projects.xls'.format(output_folder, c))

                # Increase iterator
                count += 1

                # ///////////////////// End of for loop

            end = pd.datetime.now()
            elapsed = end - start
            minutes, seconds = divmod(elapsed.total_seconds(), 60)
            f.write('\n\n')
            f.write('(Elapsed time: {:0>1} minutes {:.2f} seconds)'.format(int(minutes), seconds))

        # Delete temp files
        os.remove(os.path.join(output_folder, 'eSTIP Filtered Table.xlsx'))
        return

class RouteConverter(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = 'Convert RouteID to DisplayID'
        self.description = 'This tool is used to fill-in missing DisplayID values in the RIS_LRS enterprise geodatabase. This tool must be run with a database connection made by the LRS user, because it requires data owner priviliges\
                            to run the disenable and renable editor tracking functions. '
        self.canRunInBackground = False
        self.category = r'GIS Unit\Roads & Highways'

    def getParameterInfo(self):
        """Define parameter definitions"""
        param0 = arcpy.Parameter(
            displayName = 'Database Connection File',
            name = 'DB_connection_file',
            datatype = 'DEFile',
            parameterType = 'Required',
            direction = 'Input')
        param0.value = os.path.join(os.environ['USERPROFILE'], 'AppData\Roaming\ESRI\Desktop10.5\ArcCatalog\{}_child_1.sde'.format(os.environ['USERNAME']))
        param1 = arcpy.Parameter(
            displayName = 'Choose Output Location for Summary Text file',
            name = 'output_location',
            datatype = 'DEFolder',
            parameterType = 'Required',
            direction = 'Input')
        param1.value = desktop
        params = [param0, param1]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """ Main source of code """
        # Setup Geoprocessing Environment
        workpath = parameters[0].valueAsText
        arcpy.env.workspace = workpath

        # Compile RegEx Pattern objects only once
        RouteGroup_pattern = re.compile(r'[A-Z]{,3}')
        RouteNumber_pattern = re.compile(r'\d{1,6}(?=[PM])')   # match 1-6 digits only if followed by either uppercase'P' or 'M'
        RouteDirection_pattern = re.compile(r'(?<=\d)[PM]')    # match uppercase 'P' or 'M' only if preceded by 1-6 digits
        InterchangeLocation_pattern = re.compile(r'\d{3}\.\d') # match 3 digits followed by a '.' followed by another digit
        RampElement_pattern = re.compile(r'(?<=\.\d)[A-Z]')    # match a single uppercase letter only if preceded by a '.' and a digit

        # Define function that does RouteID -> DisplayID conversion
        def RouteConverter(RouteID):
            """ Uses regular expressions on RouteID to convert it to legacy DisplayID syntax. """

            RouteGroup = RouteGroup_pattern.search(RouteID).group(0)
            RouteNumber = RouteNumber_pattern.search(RouteID).group(0)
            RouteDirection = RouteDirection_pattern.search(RouteID).group(0)
            InterchangeLocation = InterchangeLocation_pattern.search(RouteID)
            if InterchangeLocation != None:
                InterchangeLocation = InterchangeLocation.group(0)
            RampElement = RampElement_pattern.search(RouteID)
            if RampElement != None:
                RampElement = RampElement.group(0)

            # Build DisplayID
            if InterchangeLocation == None and RampElement == None:
                DisplayID = RouteGroup + '-' + RouteNumber + '-' + RouteDirection
            elif InterchangeLocation != None and RampElement == None:
                DisplayID = RouteGroup + '-' + RouteNumber + '-' + RouteDirection + '-' + InterchangeLocation
            elif InterchangeLocation != None and RampElement != None:
                DisplayID = RouteGroup + '-' + RouteNumber + '-' + RouteDirection + '-' + InterchangeLocation + '-' + RampElement

            return DisplayID

        # List comprehension to create list of feature classes excluding feature classes known to be missing either RouteID or DisplayID (starts searching at index 12 since all fc name's have 'RIS_LRS.LRS.')
        fcList = [fc for fc in arcpy.ListFeatureClasses() if not fc.endswith(('AltStreetName', 'CountyBoundaries', 'Calibration_Point', 'Centerline', 'DistrictBoundary', 'Patrol_Boundary', 'Redline'), 12)]

        # Create context manager (with statement) and write header to RouteConverter_Log_[date]_[time] text file
        esri_dict = arcpy.GetInstallInfo()
        with open(os.path.join(desktop, 'RouteConverter_Log_{:%b_%d_%Y_%I%M}.txt'.format(today)), 'w') as f:
            f.write('Date: {: %B %d, %Y  %I:%M %p}\n'.format(today))
            f.write('ArcGIS Version: {}\n'.format(str(esri_dict['Version'])))
            f.write('Python Version: {}.{}.{}\n'.format(sys.version_info[0], sys.version_info[1], sys.version_info[2]))
            f.write('Script: RouteConverter.py\n')
            f.write('Purpose: Search all feature classes in RIS_LRS for missing DisplayID values. When it finds one,\nit uses regular expressions to parse the RouteID into components and rebuild it to DisplayID syntax.\n\n\n')
            f.write('Results:\n')

            # Setup tool progress bar & counters
            fcCount = len(fcList)
            arcpy.SetProgressor(type = 'step', message = 'Looping through feature classes in RIS_LRS SDE.', min_range = 0, max_range = fcCount, step_value = 1)
            progress_bar_counter = 1
            displayID_total_counter = 0

            # Loop through sorted list; create empty dictionary to place results: Key = feature class; Value = # of DisplayID records updated
            results = {}
            for fc in sorted(fcList):

                # Set individual fc RouteConverter counter & progress info
                fc_counter = 0
                arcpy.SetProgressorLabel('{}/{} checking {}'.format(progress_bar_counter, fcCount, fc))

                # Disable Editor Tracking (so username of person running script doesn't obliterate R&H person who had made last edit; Disable Editor Tracking function can only execute if the data owner runs the script
                #arcpy.DisableEditorTracking_management(fc)

                # List comprehension to create list of field names (function normally returns an object and here we're just interested in name attribute).
                fields = [field.name for field in arcpy.ListFields(fc)]

                # Check to see if fc has 'DisplayID' field
                if 'DisplayID' not in fields:
                    arcpy.AddMessage('{} skipped by script (missing DisplayID).'.format(fc))
                    continue   # If above statement isn't true, skip the remaining code and proceed to next fc in fcList

                # Check for RouteId and RouteID variant spellings & place into list
                RouteID_list = [field for field in fields if field == 'RouteId' or field == 'RouteID']

                # Check to see if fc has either RouteId or RouteID field.
                if len(RouteID_list) == 0:
                    arcpy.AddMessage('{} skipped by script (missing RouteID/RouteId).'.format(fc))
                    continue

                # Enterprise geodatabases require a separate edit session for each feature class
                edit = arcpy.da.Editor(workpath)
                edit.startEditing(with_undo = False, multiuser_mode = True)   # having with_undo = False gives performance boost according to ESRI (don't need rollback capability)
                edit.startOperation()

                # For feature classes, with RouteId spelling (e.g. RIS_LRS.LRS.Routes), update the DisplayID if the value is NULL or empty string.
                if RouteID_list[0] == 'RouteId':
                    with arcpy.da.UpdateCursor(fc, field_names = ['OID@', 'RouteId', 'DisplayID'], where_clause = " DisplayID IS NULL OR DisplayID = '' ", sql_clause = (None, " ORDER BY RouteID ")) as cursor:
                        for row in cursor:
                            RouteID = row[1]
                            row[2] = RouteConverter(RouteID)
                            #cursor.updateRow(row)
                            fc_counter += 1
                            displayID_total_counter += 1
                            f.write('{}  OBJECTID:{}  {} converted to {}\n'.format(fc, str(row[0]), row[1], row[2]))

                # And now for feature classes, with RouteID spelling
                elif RouteID_list[0] == 'RouteID':
                    with arcpy.da.UpdateCursor(fc, field_names = ['OID@', 'RouteID', 'DisplayID'], where_clause = " DisplayID IS NULL OR DisplayID = '' ", sql_clause = (None, " ORDER BY RouteID ")) as cursor:                   # RouteID IN ('FL5119P', 'FL5237P', 'FL5237P', 'FL5513P', 'FL5668P', 'FL5674P', 'FL5857P', 'FL5990P', 'FR1015P', 'FR2045P', 'NM211P', 'US285P', 'US491P', 'US550P', 'US64P', 'US70P', 'US84P')
                        for row in cursor:
                            RouteID = row[1]
                            row[2] = RouteConverter(RouteID)
                            #cursor.updateRow(row)
                            fc_counter += 1
                            displayID_total_counter += 1
                            f.write('{}  OBJECTID: {}  {} converted to {}\n'.format(fc, str(row[0]), row[1], row[2]))

                # Close and save edit session within Enterprise geodatabase
                edit.stopOperation()
                edit.stopEditing(save_changes = True)

                if fc_counter > 0:
                    f.write('{} summary: {} DisplayId records were updated\n\n\n'.format(fc, fc_counter))
                    arcpy.AddMessage('{} summary: {} DisplayId records were updated'.format(fc, fc_counter))

                # Reenable Editor Tracking on the fc (Enable Editor Tracking function can only execute if the data owner runs the script)
                #arcpy.EnableEditorTracking_management(fc, 'created_user', 'created_date', 'last_edited_user', 'last_edited_date', add_fields = 'NO_ADD_FIELDS')

                # Add results to dictionary (ignore the unicode fc names and convert them to ascii; also ignore the 'RIS_LRS.LRS.' prefix present on all feature classes):
                results[fc[12:].encode('ascii', 'ignore')] = fc_counter

                # Increment blue progress bar, as well as current fc number that tool is at in for loop
                arcpy.SetProgressorPosition()
                progress_bar_counter += 1

            # Summarize total DisplayIDs updated
            f.write('RouteConverter script summary: {} DisplayId records were updated.'.format(displayID_total_counter))
            arcpy.AddMessage('\nRouteConverter script summary: {} DisplayId records were updated\n'.format(displayID_total_counter))
            #f.write('{}'.format(results))

        return