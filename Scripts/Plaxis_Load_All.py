import Shovel_Classes
import sys
import toolkit

if __name__ == '__main__':
    try:
        wb_name = sys.argv[1]
        load = Shovel_Classes.Loader(shovel_wb=wb_name, load_all=True)
        load.extract_to_table()
    except Exception as e:
        toolkit.error_occur(e)