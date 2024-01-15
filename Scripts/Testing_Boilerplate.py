from Shovel_Classes import Boilerplate

def testing_check_plaxis(bp: Boilerplate):
    print(bp.app_check_plaxis(terminate=True))

def testing_plaxis_launcher(bp: Boilerplate):
    proc, s, g = bp.app_plaxis_launcher(plx_output=True)
    print(proc)
    print(s.name)

def testing_open_model(bp: Boilerplate, model_path: str):
    bp.app_plaxis_launcher(plx_output=True)
    bp.plx_open(model_path)

bp = Boilerplate(10000, r"VakNTMK/=W~>@4~+", r"C:\Program Files\Bentley\Geotechnical\PLAXIS 2D CONNECT Edition V21")
# testing_check_plaxis(bp)
testing_plaxis_launcher(bp)
testing_open_model(bp, r"C:\Users\Sher Wen\Desktop\CR113 Wei Ern\Section D\Plaxis V2\CR113_D_CaseA_SLS_0.7EI_SpD_02pit.p2dx")
