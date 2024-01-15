import win32com.client
import pandas as pd
import numpy as np
import psutil as psu
import plxscripting.easy as plx
import subprocess
import time
import toolkit

### Notes ###

# Extraction profile tuples follow the following format: Phase, Elements, Property, Envelope.
# ExtractionProfile objects will be created per model, data will also be written into excel per model

# Subprofiles follow the following format: Phase, Element, property
# Each model will have a series of unique subprofiles to optimize extraction
# ExtractionProfile objects will have subprofiles depending on their envelope

class ExcelTable:
    """ This class converts a given excel ListObject into a pandas Dataframe
        and also writes data from dataframe to ListObject
        #TESTED 14/2/2023
    """
    def __init__(self, wb, sheet, tablename):
        """Creates a listobject from the table specified
        """
        xl = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.wb = xl.Workbooks(wb)
        self.sheet = self.wb.Worksheets(sheet)
        self.table = self.sheet.ListObjects(tablename)
        self.df = None

    @classmethod
    def open_wb(cls, wb_path, sheet, tablename):
        """ to be used to initiate a class instance when the workbook is not opened
        """
        xl = win32com.client.gencache.EnsureDispatch('Excel.Application')
        wb = xl.Workbooks.Open(wb_path)
        wb_name = str(wb.Name.split(".")[0])
        return cls(wb_name, sheet, tablename)

    def headerlist(self):
        """Retrieves the header row range as a Tuple
        """
        return self.table.HeaderRowRange()[0]
    
    def dataframe(self, first_col_as_index: bool = False) -> object:
        """ Generates a dataframe based on the ListObject. First Column can optionally be the index.
        """
        tbl_values = self.table.Range.Value
        self.df = pd.DataFrame(tbl_values[1:],columns=tbl_values[0])
        if first_col_as_index:
            self.df.set_index(self.df.iloc[:, 0].name, inplace=True)
        return self.df
    
    def df_to_dict(self, column_key: str, column_values: list = [], filter_column: str = "", filter_data: str = "") -> dict:
        """ Generates a dictionary with a 1 column value as key and multiple columns as list
            1 optional filter can be applied to get the corresponding values
            If there is only one column value, will return key : value instead of [values]
        """
        if filter_column and filter_data:
            df1 = self.df.where(self.df[filter_column] == filter_data).dropna()
        else:
            df1 = self.df

        df1.reset_index(inplace= True)
        if column_values == []:
            df2 = df1
        else:
            df2 = df1[[column_key] + column_values]

        if len(df2.columns) == 2:
            return dict(df2.values)
        return df2.set_index(column_key).T.to_dict('list')

    def clear_table(self):
        """ Clears a table object and reset number of rows to 1
        """
        self.table.DataBodyRange.Delete()
        return
    
    def search_and_delete(self, delete_dict: dict):
        """ Dictionary of column name:[items to be removed]
        """
        if delete_dict == False:
            return

        for column in delete_dict:
            for item in delete_dict[column]:
                try:
                    self.table.Range.AutoFilter(self.table.ListColumns(column).Index, item)
                    delete_range = self.table.DataBodyRange.SpecialCells(12) # enum: xlAllVisibleCells
                    self.table.AutoFilter.ShowAllData()
                    delete_range.Delete(-4162) # enum: xlShiftUp'
                except:
                    if self.table.ShowAutoFilter:
                        if self.table.AutoFilter.FilterMode:
                            self.table.AutoFilter.ShowAllData()
                    pass
        return
    
    def write_dict_to_table(self, data_dict: dict, wipe_table: bool = False, start_col: int = 1):
        """ Writes a dictionary consisting of data:iterable data pair into excel table
            the Value iterable will be unpacked into every row, while the
            key will be repeated for every row in its respective columns

            Value iterable can 1 list or a list of lists for unpacking into its respective columns
            this function assumes that every list is of the same length
        """
        if wipe_table:
            self.clear_table()
        
        for key in data_dict:
            data = data_dict[key]

            if not data:
                input_data = None
                data_col_count = 0
                data_row_count = 1
            elif isinstance(data[0], list):
                input_data = [list(d) for d in zip(*data)]
                data_col_count = len(input_data[0])
                data_row_count = len(input_data)
            else:
                input_data = [[d] for d in data]
                data_col_count = 1
                data_row_count = len(input_data)

            if not isinstance(key, tuple): #key has to be hashable type
                key = (key,)
            key_col_count = len(key)
            
            if (key_col_count + data_col_count > self.table.ListColumns.Count - start_col + 1):
                raise Exception("Data will fall outside table")
            else:
                last_row = self.table.ListRows.Count + 1

                key_start_range = self.table.ListColumns(start_col).Range(last_row, 1).GetOffset(1, 0)
                key_end_range = key_start_range.GetOffset(data_row_count - 1, key_col_count - 1)
                if input_data:
                    data_start_range = self.table.ListColumns(start_col).Range(last_row, 1).GetOffset(1, key_col_count)
                    data_end_range = data_start_range.GetOffset(data_row_count - 1, data_col_count - 1)

                self.sheet.Range(key_start_range, key_end_range).Value = key
                if input_data:
                    self.sheet.Range(data_start_range, data_end_range).Value = input_data


    def write_df_to_table(self, df: object, wipe_table: bool = False, starting_col: int = 1) -> bool:
        """ Inserts a dataframe to the bottom of a table. Option given to wipe table first. Index column will be excluded
        """       
        if wipe_table:
            self.clear_table()

        # Converts pd dataframe into numpy C-Contiguous array for writing into excel with wincom32
        # Then converts to list. Index column will be excluded
        np_array = np.ascontiguousarray(df)
        data_list = np_array.tolist()


        last_row = self.table.ListRows.Count + 1
        # Getting entire pasting range
        start_range = self.table.ListColumns(starting_col).Range(last_row, 1).GetOffset(1, 0)
        end_range = start_range.GetOffset(len(df.index) - 1, len(df.columns) - 1)
        self.sheet.Range(start_range, end_range).Value = data_list # Paste data
        return True

    def close(self, save: bool=False):
        self.wb.Close(save)

    def save_as(self, full_path, close: bool=False):
        self.wb.SaveAs(full_path)
        if close:
            self.wb.Close(False)
            
class Boilerplate:
    """ This class represents ONE application instance of either Plaxis Input or Output.
        If both Input and Output must be open, another instance of this class must be created.

    """
    def __init__(self, host: str, port: int, password: str, plaxis_folder: str):
        self.host = host
        self.port = int(port)
        self.password = password
        self.s = None
        self.g = None
        self.process = None
        if plaxis_folder[-1] == "\\":
            self.plaxis_folder = plaxis_folder
        else:
            self.plaxis_folder = plaxis_folder + "\\"

    def app_check_plaxis(self, plx_output: bool = False, terminate: bool = False) -> object:
        """ Checks if Plaxis is currently running. Default launches plaxis input.
            Enable terminate to ensure that process is killed. Will return process object otherwise if found.
        """
        if plx_output:
            app_name = "Plaxis2DOutput.exe"
        else:
            app_name = "Plaxis2DXInput.exe"

        for process in psu.process_iter():
            if process.name() == app_name:
                if terminate:
                    process.terminate()
                    return None
                else:
                    return process
        return None

    def app_plaxis_launcher(self, plx_output: bool = False, timeout: float = 5.0) -> tuple:
        """ Launches Plaxis with optional timeout setting, recommended to stick to 5.0
            Default launches plaxis input.
        """
        if plx_output:
            app_name = "Plaxis2DOutput.exe"
        else:
            app_name = "Plaxis2DXInput.exe"

        args = [self.plaxis_folder + app_name, f"--AppServerPort={self.port}", f"--AppServerPassWord={self.password}"]
        self.process = subprocess.Popen(args)
        self.s, self.g = plx.new_server(address=self.host, port=self.port, timeout=timeout, password=self.password)
        if self.s.active == True:
            return(self.process, self.s, self.g)
        return None

class Loader:
    """ This class iterates through the model table, opens models that are queued to be loaded
        and extracts all element and phase data from them
    """
    def __init__(self, shovel_wb: str, load_all: bool=False):
        settings = ExcelTable(shovel_wb,"Plaxis_extractor", "tbl_Settings")
        settings.dataframe()
        self.settings_dict = settings.df_to_dict("Settings", ["Value"])
        self.bp = Boilerplate(self.settings_dict["Host"], self.settings_dict["Plaxis output port"], 
                              self.settings_dict["Plaxis password"], self.settings_dict["Plaxis installation folder"])
        self.bp.app_check_plaxis(plx_output=True, terminate=True)
        process, self.s, self.g = self.bp.app_plaxis_launcher(plx_output=True)

        self.element_tbl = ExcelTable(shovel_wb, "_system", "tbl_AllElements")
        model_tbl = ExcelTable(shovel_wb, "Plaxis_extractor", "tbl_PlaxisFiles")
        model_tbl.dataframe()
        self.load_all = load_all
        if load_all:
            self.model_dict = model_tbl.df_to_dict("Model Name", ["Path"])
        else:
            self.model_dict = model_tbl.df_to_dict("Model Name", ["Path"], "Action", "Load Model")

    def extract_to_table(self) -> bool:
        """ Extracts all element and phase data of the current model into shovel AllElements table
        """
        if not self.model_dict:
            toolkit.mbox("Loading model info into excel...", "No model to load")
            exit()

        if self.load_all:
            try:
                self.element_tbl.clear_table()
            except Exception:
                pass

        window = toolkit.create_window("Loading model info into excel...") # Window for progressbar
        extraction_list = ["phases", "plates", "EmbeddedBeamRows", "NodeToNodeAnchors", "FixedEndAnchors", "Geogrids", "Interfaces"]
        new_elements_dict = {}
        for i, model in enumerate(self.model_dict, start=1):
            progressbar = toolkit.create_progressbar(window, model, i) # 1 bar per model
            path = self.model_dict[model]
            if not self.load_all:
                self.element_tbl.search_and_delete({"Model": [model]}) 
            self.s.open(path)

            for elem_type in extraction_list:
                try:
                    if elem_type != "phases":
                        new_elements_dict[model] = new_elements_dict.get(model, []) + [element.Name.value for element in eval("self.g." + elem_type)]
                    else:
                        new_elements_dict[model] = new_elements_dict.get(model, []) + [element.Identification.value for element in eval("self.g." + elem_type)]
                except:
                    pass

                toolkit.step_progressbar(progressbar, 14.3)

        self.element_tbl.write_dict_to_table(new_elements_dict, start_col=2)
        window.destroy()
        toolkit.mbox("Plaxis model extraction", "Done!")

class Extractor:

    def __init__(self, shovel_wb: str, extract_all: bool=False) -> None:
        settings = ExcelTable(shovel_wb,"Plaxis_extractor", "tbl_Settings")
        settings.dataframe()
        self.settings_dict = settings.df_to_dict("Settings", ["Value"])

        # Plaxis output boilerplate object
        self.bp_output = Boilerplate(self.settings_dict["Host"], self.settings_dict["Plaxis output port"], 
                              self.settings_dict["Plaxis password"], self.settings_dict["Plaxis installation folder"])
        self.bp_output.app_check_plaxis(plx_output=True, terminate=True)
        self.process, self.s_o, self.g_o = self.bp_output.app_plaxis_launcher(plx_output=True)
        
        # Extraction table converted to dataframe
        extraction_tbl = ExcelTable(shovel_wb, "Plaxis_extractor", "tbl_Extraction")
        self.extraction_df = extraction_tbl.dataframe().drop(columns=["Element Type"])
        profiles_tbl = ExcelTable(shovel_wb, "Plaxis_extractor", "tbl_Profiles")
        self.profiles_df = profiles_tbl.dataframe(first_col_as_index=True)
        extraction_models = set(self.extraction_df.loc[:, 'Model'].values.tolist())

        # Get dictionary of models to be extracted
        model_tbl = ExcelTable(shovel_wb, "Plaxis_extractor", "tbl_PlaxisFiles")
        model_tbl.dataframe()
        self.extract_all = extract_all
        if extract_all:
            all_model_dict = model_tbl.df_to_dict("Model Name", ["Path"])
        else:
            all_model_dict = model_tbl.df_to_dict("Model Name", ["Path"], "Action", "Extract Data")
        self.model_dict = {model:all_model_dict[model] for model in extraction_models}

        # Class variables to be used in later functions
        self.all_profiles = []
        self.output_data_table = None
        self.output_map_table = None
        self.output_path = self.settings_dict["Output folder path"]
        self.project_name = self.settings_dict["Project Name"]
        self.output_max_index = 0
        self.output_istemplate = True
        self.window = toolkit.create_window("Plaxis data extraction")
        self.progressbar = None


    def plx_open_model(self, model_path: str):
        self.s_o.open(model_path)
        return
    
    def plx_extract_model(self, phase_str: str, elem_str: str, property: str) -> list: 
        """ Extracts data from plaxis based on subprofile list

        """
        g = self.g_o
        
        try:
            elem_type = elem_str.split('_')[0]
            if elem_type == 'NegativeInterface' or elem_type == 'PositiveInterface':
                elem_type = 'Interface'
            phase_ID = phase_str[phase_str.find('[')+1:phase_str.find(']')]
            phase_obj = eval(f"g.{phase_ID}")
            element_obj = eval(f"g.{elem_str}")
            resulttype_obj = eval(f"g.ResultTypes.{elem_type}.{property}")
            results = g.getresults(element_obj, phase_obj, resulttype_obj, 'node')
            data = [r for r in results]
        except:
            data = []         
        return data

    def xl_open_extraction_wb(self):
        """ Opens extraction workbook defined in the shovel settings table,
            If no path is provided, use the template workbook specified.
        """
        if self.settings_dict["Existing Excel path"] == None:
            if self.settings_dict["Template Excel path"] == None:
                toolkit.mbox("Extraction stopped", "Please enter path of extraction file or template")
                exit()
            else:
                wb_path = self.settings_dict["Template Excel path"]
                self.output_istemplate = True
        else:
            wb_path = self.settings_dict["Existing Excel path"]
            self.output_istemplate = False

        self.output_data_table = ExcelTable.open_wb(wb_path, "Extractor", "tbl_Data")
        wb_name = self.output_data_table.wb.Name
        self.output_map_table = ExcelTable(wb_name, "Extractor", "tbl_Extraction")
        self.output_profile_table = ExcelTable(wb_name, "_Profiles", "tbl_AllProfiles")

        if self.extract_all:
            self.output_data_table.clear_table()
            self.output_map_table.clear_table() 
        return
    
    def xl_remove_profiles(self, model_name: str):
        """ Removes all profiles of the current model in the output table. Sets the max_index value for add profile step
        """
        
        if (self.output_map_table.table.ListRows.Count <= 1) and (self.output_data_table.table.ListRows.Count <= 1):
            self.output_max_index = 0
        elif self.extract_all or self.output_istemplate:
            self.output_max_index = int(max(self.output_map_table.table.ListColumns("Index").DataBodyRange())[0])
        else:
            self.output_map_table.dataframe()
            current_dict = self.output_map_table.df_to_dict("Index", filter_column="Model", filter_data=model_name)
            self.output_map_table.search_and_delete({"Model":model_name})
            for index in current_dict:
                self.output_data_table.search_and_delete({"Index": index})
            self.output_max_index = int(max(self.output_map_table.table.ListColumns("Index").DataBodyRange())[0])
        return

    def xl_add_profiles(self, model_name: str):
        """ Adds data to the output table
        """
        # Extraction list
        extractions = self.extraction_df.loc[self.extraction_df['Model'] == model_name, ~self.extraction_df.columns.isin(['Model', 'Element Type'])].values.tolist()
        clean_ext = []
        for e in extractions: #Remove None
            cleaned = [a for a in e if a is not None]
            if len(cleaned):
                clean_ext.append(cleaned)

        required_profiles = set([ext[-1] for ext in clean_ext])
        profile_list = self.profiles_df.loc[self.profiles_df['Name'].isin(required_profiles), ['Name'] + list(self.profiles_df.loc[:,'Property 1':'Property 8'])].values.tolist()
        profile_dict = {profile[0]:[p for p in profile[1:] if p is not None] for profile in profile_list} # Remove None

        sorted_extractions = sorted(clean_ext, key=lambda a : a[1]) #sort by phase
        step = 100/len(sorted_extractions)

        for ext in sorted_extractions:
            no = ext[0]
            phase = ext[1]
            elements = ext[2:-1]
            profile = ext[-1]

            for element in elements:
                x_coord = self.plx_extract_model(phase, element, 'X')
                y_coord = self.plx_extract_model(phase, element, 'Y')
                data_list = []
                properties = profile_dict[profile]
                for prop in properties: # iterates through the properties in the profile indicated
                    data_list.append(self.plx_extract_model(phase, element, prop))

                self.output_max_index += 1
                if not data_list[0]:
                    status = 'No data'
                else:
                    status = 'Extracted'
                    new_data = {self.output_max_index: [x_coord, y_coord, *data_list]} # dictionary of index: [x coord, y coord, data]
                    self.output_data_table.write_dict_to_table(new_data) 

                extraction_dict = {(self.output_max_index, no, model_name, phase, element, profile, status): []} # dictionary of profile: empty list
                self.output_map_table.write_dict_to_table(extraction_dict)

            toolkit.step_progressbar(self.progressbar, step)
        
        self.output_profile_table.write_df_to_table(self.profiles_df, wipe_table=True) # Copies entire profile table over

    def xl_save_wb(self):
        wb = self.output_map_table.wb
        if self.output_istemplate:
            new_file_name = self.project_name + ' Plaxis Extraction'
            wb.SaveAs(f'{self.output_path}\\{new_file_name}', 52) # enum for .xlsm format
            wb.Activate()
        else:
            wb.Save()
            wb.Activate()

    def process_flow(self):
        if not self.model_dict:
            toolkit.mbox("Plaxis Extraction", "No models to extract from")
            return
        
        self.xl_open_extraction_wb()
        for i, model in enumerate(self.model_dict, start=1):
            self.progressbar = toolkit.create_progressbar(self.window, model, i)

            path = self.model_dict[model]
            self.plx_open_model(path)
            self.xl_remove_profiles(model)
            self.xl_add_profiles(model)

        self.xl_save_wb()
        self.process.terminate()
        self.window.destroy()
        toolkit.mbox("Plaxis Extraction", "Extraction complete")