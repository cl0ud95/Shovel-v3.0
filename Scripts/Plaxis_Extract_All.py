import Shovel_Classes
import sys
import toolkit

if __name__ == '__main__':
    try:
        wb_name = sys.argv[1]
        extract = Shovel_Classes.Extractor(shovel_wb=wb_name, extract_all=True)
        extract.process_flow()
    except Exception as e:
        toolkit.error_occur(e)

# wb_name = "Shovel.xlsm"
# extract = Shovel_Classes.Extractor(shovel_wb=wb_name, extract_all=True)
# extract.process_flow()