# -*- coding: utf-8 -*-
import win32com.client as win32
import os
import pandas as pd
import math

# 專案設定
# 這裡請替換為你要開啟的模型檔案路徑
MODEL_PATH = r"D:\Users\63427\Desktop\Code\EQ2SAP\example\model_test.sdb"


class Sap2000(object):
    def __init__(self):
        self.SapObject = None
        self.SapModel = None

    def initializeNewModel(self, unitsTag=12):
        """
        初始化一個新的 Sap2000 模型。
        """
        self.SapObject = win32.Dispatch("SAP2000.SapObject")  # create SAP2000 object
        self.SapObject.ApplicationStart()  # start a SAP2000 program
        self.SapModel = self.SapObject.SapModel  # create SAP2000 model object
        self.SapModel.InitializeNewModel(
            unitsTag
        )  # Clears the previous model and initializes a new model

    def file_OpenFile(self, FileName):
        """
        開啟現有的 Sap2000 模型檔案。
        支援副檔名：.sdb（標準 Sap2000 檔案）、$2k/.s2k（文字檔）、.xlsx/.xls（Excel 檔案）、.mdb（Access 檔案）。

        參數：
            FileName (str): 欲於 Sap2000 開啟的模型檔案完整路徑。
        """
        self.SapModel.File.OpenFile(FileName)

    def file_Save(self, FileName):
        """
        儲存目前 Sap2000 模型檔案。

        參數：
            FileName (str): 儲存檔案的完整路徑，建議使用 .sdb 副檔名。
                若未指定檔名，則以目前檔名儲存。若模型尚未儲存過且未指定檔名，將回傳錯誤。

        回傳：
            int: 儲存成功回傳 0，否則回傳非 0。
        """
        self.SapModel.File.Save(FileName)  # eg."C:\SapAPI\x.sdb"

    def closeModel(self):
        """
        ---close SAP2000 model---
        """
        # close SAP2000 model
        self.SapObject.ApplicationExit(True) #True means save the model before close,False otherwise.
        self.SapModel=0 # release the memory
        self.SapObject=0 # release the memory

    def setUnits(self, unitsTag):
        """
        設定目前 Sap2000 模型的單位。

        參數：
            unitsTag (int): 單位代碼。
                1=lb_in_F, 2=lb_ft_F, 3=kip_in_F, 4=kip_ft_F, 5=kN_mm_C, 6=kN_m_C,
                7=kgf_mm_C, 8=kgf_m_C, 9=N_mm_C, 10=N_m_C, 11=Ton_mm_C, 12=Ton_m_C,
                13=kN_cm_C, 14=kgf_cm_C, 15=N_cm_C, 16=Ton_cm_C。
        """
        self.SapModel.SetPresentUnits(unitsTag)

    def groupdef_getnamelist(self):
        """
        取得目前 Sap2000 模型的群組名稱列表。

        Returns:
            bool: 讀取是否成功。
            int: 群組數量。
            list: 群組名稱（list of str）。
        """
        ret = self.SapModel.GroupDef.GetNameList()
        return ret

    def loadcases_getnamelist(self):
        """
        取得目前 Sap2000 模型的載重名稱列表。

        Returns:
            bool: 讀取是否成功。
            int: 群組數量。
            list: 群組名稱（list of str）。
        """
        ret = self.SapModel.LoadCases.GetNameList()
        return ret

    def define_LoadCases_StaticLinear_SetCase(self, name):
        """
        初始化一個靜力線性載重案例。

        參數：
            name (str): 載重案例名稱（可為新建或既有）。
        """
        self.SapModel.LoadCases.StaticLinear.SetCase(name)

    def define_LoadCases_StaticLinear_SetLoads(
        self, name, numberLoads, loadType, loadName, scaleFactor
    ):
        """
        設定指定分析案例的載重資料。

        參數：
            name (str): 既有靜力線性載重案例名稱。
            numberLoads (int): 指定案例的載重數量。
            loadType (list of str): 載重型態（'Load' 或 'Accel'）。
            loadName (list of str): 載重名稱。若 loadType 為 'Load'，此為已定義載重名稱；若為 'Accel'，此為 UX、UY、UZ、RX、RY 或 RZ。
            scaleFactor (list of float): 各載重的比例因子。對於 Accel UX/UY/UZ 單位為 L/s²，其餘無單位。
        """
        self.SapModel.LoadCases.StaticLinear.SetLoads(
            name, numberLoads, loadType, loadName, scaleFactor
        )

    def analyze_SetRunCaseFlag(self, Name, Run, All=False):
        """
        設定載重案例的執行旗標。

        參數：
            Name (str): 欲設定執行旗標的載重案例名稱。
            Run (bool): 是否執行該載重案例。
            All (bool): 若為 True，則所有載重案例皆依 Run 設定，忽略 Name。
        """
        self.SapModel.Analyze.SetRunCaseFlag(Name, Run, All)

    def analyze_RunAnalysis(self):
        """
        執行模型分析。

        注意：模型必須已儲存（有檔案路徑）才能進行分析。若為新建模型，請先呼叫 File.Save。
        成功時回傳 0，否則回傳非 0。
        """
        ret = self.SapModel.Analyze.RunAnalysis()
        return ret

    def results_JointDispl(self, Name, ItemTypeElm=0):
        """
        回傳指定點元素的相對位移結果。

        參數
        ------
        Name : str
            現有點物件、點元素或群組名稱，依 ItemTypeElm 決定。
        ItemTypeElm : int, 預設 0
            查詢型態：
                0 - ObjectElm：依點物件名稱
                1 - Element：依點元素名稱
                2 - GroupElm：群組內所有點元素
                3 - SelectionElm：所有已選取點元素（忽略 Name）

        回傳
        ------
        tuple
            (index, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
            NumberResults : int
                結果總數。
            Obj : list of str
                各結果對應的點物件名稱（可能為空字串）。
            Elm : list of str
                各結果對應的點元素名稱。
            LoadCase : list of str
                各結果對應的分析案例或組合名稱。
            StepType : list of str
                各結果的步驟型態。
            StepNum : list of int
                各結果的步驟編號。
            U1, U2, U3 : list of float
                各結果於局部 1、2、3 軸方向的位移 [長度]。
            R1, R2, R3 : list of float
                各結果於局部 1、2、3 軸的轉角 [弧度]。
        """
        result = self.SapModel.Results.JointDispl(Name, ItemTypeElm)
        return result

    def results_AssembledJointMass(self, Name, itemTypeElm):
        """
        回傳指定點元素的組裝質量資訊。
        Note: V19之後SAP才有AssembledJointMass_1，語法不同。

        參數
        ------
        Name : str
            現有點元素或群組名稱，依 itemTypeElm 決定。
        itemTypeElm : int
            查詢型態：
                0 - ObjectElm：依點物件名稱
                1 - Element：依點元素名稱
                2 - GroupElm：群組內所有點元素
                3 - SelectionElm：所有已選取點元素（忽略 Name）

        回傳
        ------
        tuple
            (index, NumberResults, PointElm, MassSource, U1, U2, U3, R1, R2, R3)
            NumberResults : int
                結果總數。
            PointElm : list of str
                各結果對應的點元素名稱。
            U1, U2, U3 : list of float
                各結果於局部 1、2、3 軸方向的平移質量 [質量]。
            R1, R2, R3 : list of float
                各結果於局部 1、2、3 軸的轉動慣量 [質量×長度^2]。
        """
        result = self.SapModel.Results.AssembledJointMass(Name, itemTypeElm)
        return result
    
    def results_Setup_DeselectAllCasesAndCombosForOutput(self):
        """
        取消所有載重案例與組合的輸出選取。
        """
        self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

    def results_Setup_SetCaseSelectedForOutput(self,Name,Selected=True):
        """
        設定指定載重案例是否選取為輸出。

        參數：
            Name (str): 載重案例名稱。
            Selected (bool): 是否選取為輸出，預設為 True。
        """
        self.SapModel.Results.Setup.SetCaseSelectedForOutput(Name,Selected)


def get_disp(sapobj, lc_dir, gp_list, disp_note):
    jointdisp = {}
    sapobj.results_Setup_DeselectAllCasesAndCombosForOutput()
    sapobj.results_Setup_SetCaseSelectedForOutput(lc_dir, Selected=True)

    for lcg in gp_list:
        res = None
        res = sapobj.results_JointDispl(lcg, ItemTypeElm=2)
        jointdisp[lcg] = {}
        jointdisp[lcg]['node_num'] = res[1]
        jointdisp[lcg]['node_name'] = res[3]
        jointdisp[lcg]['node_disp'] = res[disp_note]
        
    return jointdisp

def get_mass(sapobj, gp_list, disp_note):
    jointmass = {}
    sapobj.results_Setup_DeselectAllCasesAndCombosForOutput()

    for gp in gp_list:
        res = sapobj.results_AssembledJointMass(gp, 2)
        jointmass[gp] = {}
        jointmass[gp]['node_num'] = res[1]
        jointmass[gp]['node_name'] = res[2]
        jointmass[gp]['node_mass'] = res[disp_note]
        
    return jointmass

def cal_period(jointdisp, jointmass, group):
    g = 9.81  # 重力加速度，與前面設定的 'Accel' 載重一致
    dict_period = {}
    for gp in group:
        dict_disp = jointdisp[gp]
        dict_mass = jointmass[gp]

        # 依據節點名稱('node_name')將位移與質量進行配對。
        # 這樣可以確保即使節點順序不同也能正確匹配。
        disp_by_node = dict(zip(dict_disp['node_name'], dict_disp['node_disp']))
        mass_by_node = dict(zip(dict_mass['node_name'], dict_mass['node_mass']))

        # 針對共通節點計算 位移 * 質量
        wu = {node: disp_by_node[node] * mass_by_node[node] for node in disp_by_node.keys() & mass_by_node.keys()}
        # 針對共通節點計算 位移 * 位移 * 質量
        wuu = {node: disp_by_node[node] * disp_by_node[node] * mass_by_node[node] for node in disp_by_node.keys() & mass_by_node.keys()}

        beta = abs(sum(wu.values()))
        zeta = sum(wuu.values())

        # 根據 Rayleigh's method 計算週期
        # T = 2 * pi * sqrt( (sum(m*u^2)) / (g * sum(m*u)) )
        period = 2 * math.pi * math.sqrt(zeta / (g * beta))

        dict_period[gp] = {}
        dict_period[gp]['period'] = period
        dict_period[gp]['beta'] = beta
        dict_period[gp]['zeta'] = zeta
        dict_period[gp]['mass'] = mass_by_node
        dict_period[gp]['disp'] = disp_by_node

    return dict_period

def export_results_to_excel(period_x, period_y, period_z, output_path):
    """
    將計算結果彙整並輸出至包含多個工作表的 Excel 檔案。

    - 'Period Calculation': 週期計算總覽。
    - 'Group-Direction': 各群組與方向的詳細節點質量與位移。

    參數:
        period_x (dict): X 方向的週期計算結果。
        period_y (dict): Y 方向的週期計算結果。
        period_z (dict): Z 方向的週期計算結果。
        output_path (str): Excel 檔案的完整輸出路徑。
    """
    # 1. 彙整總覽結果
    all_results = []
    for direction, period_data in [('X', period_x), ('Y', period_y), ('Z', period_z)]:
        for group_name, data in period_data.items():
            all_results.append({
                'Group': group_name,
                'Direction': direction,
                'Period (s)': data['period'],
                'Sum(wu)': data['beta'],
                'Sum(wuu)': data['zeta']
            })
    df_results = pd.DataFrame(all_results)

    # 2. 使用 pd.ExcelWriter 寫入多個工作表
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 2a. 寫入週期計算總表
        df_results.to_excel(writer, index=False, sheet_name='Period Calculation', float_format="%.4f")

        # 2b. 遍歷每個方向和群組，寫入詳細的 disp 和 mass 工作表
        for direction, period_data in [('X', period_x), ('Y', period_y), ('Z', period_z)]:
            for group_name, data in period_data.items():
                mass_dict = data.get('mass', {})
                disp_dict = data.get('disp', {})
                
                all_nodes = sorted(
                    list(mass_dict.keys() | disp_dict.keys()), 
                    key=lambda x: (0, int(x)) if x.isdigit() else (1, x)
                )

                detail_data = []
                for node in all_nodes:
                    detail_data.append({
                        'Node': node,
                        'Mass': mass_dict.get(node),
                        'Displacement': disp_dict.get(node)
                    })
                
                if not detail_data:
                    continue

                df_details = pd.DataFrame(detail_data)
                sheet_name = f"{group_name}-{direction}"[:31]
                df_details.to_excel(writer, index=False, sheet_name=sheet_name, float_format="%.6e")

    print(f"[訊息]：計算結果已成功匯出至：{output_path}")

def run_analysis_and_export_results(model_path, groups_x, groups_y, groups_z):
    """
    執行完整的 SAP2000 週期分析與結果匯出流程。

    1. 開啟並準備 SAP2000 模型。
    2. 設定並執行單位力載重分析。
    3. 提取位移與質量，計算各群組與方向的週期。
    4. 將結果印至主控台並匯出至 Excel 檔案。

    參數:
        model_path (str): SAP2000 模型檔案的完整路徑。
        groups_x (list): 要在 X 方向分析的群組名稱列表。
        groups_y (list): 要在 Y 方向分析的群組名稱列表。
        groups_z (list): 要在 Z 方向分析的群組名稱列表。
    """
    # --- 1. 模型物件創建及控制 ---
    sapmodel = Sap2000()
    sapmodel.initializeNewModel()

    if not os.path.exists(model_path):
        print(f"[錯誤]：找不到模型檔案 -> {model_path}")
        exit(1)

    print(f"[訊息]：正在開啟模型檔案：{model_path}...")
    sapmodel.file_OpenFile(model_path)
    print("[訊息]：模型已成功開啟！")
    sapmodel.setUnits(12)

    # --- 2. 分析用力量加載及計算 ---
    lc_unit = {"UNIT-X": "UX", "UNIT-Y": "UY", "UNIT-Z": "UZ"}
    for lc, dir_name in lc_unit.items():
        sapmodel.define_LoadCases_StaticLinear_SetCase(lc)
        sapmodel.define_LoadCases_StaticLinear_SetLoads(
            lc, 1, ["Accel"], [dir_name], [9.81]
        )
    print("[訊息]：單位均佈力載重設定完成！")

    _, num_lc, namelist_lc = sapmodel.loadcases_getnamelist()
    namelist_lc = list(namelist_lc)
    eqlc_list = list(lc_unit.keys())
    sapmodel.file_Save(model_path)
    for lc in namelist_lc:
        if lc in eqlc_list:
            sapmodel.analyze_SetRunCaseFlag(lc, True)
        else:
            sapmodel.analyze_SetRunCaseFlag(lc, False)

    runstatus = sapmodel.analyze_RunAnalysis()
    if runstatus != 0:
        print("[警告]：模型分析未成功執行。")
    else:
        print("[訊息]：模型分析完成！")

    # --- 3. 獲取分析結果並計算週期 ---
    jointdisp_x = get_disp(sapmodel, 'UNIT-X', groups_x, 7)
    jointdisp_y = get_disp(sapmodel, 'UNIT-Y', groups_y, 8)
    jointdisp_z = get_disp(sapmodel, 'UNIT-Z', groups_z, 9)

    jointmass_x = get_mass(sapmodel, groups_x, 3)
    jointmass_y = get_mass(sapmodel, groups_y, 4)
    jointmass_z = get_mass(sapmodel, groups_z, 5)

    period_x = cal_period(jointdisp_x, jointmass_x, groups_x)
    period_y = cal_period(jointdisp_y, jointmass_y, groups_y)
    period_z = cal_period(jointdisp_z, jointmass_z, groups_z)
    
    print("[訊息]：週期計算完成！")

    # 關閉 SAP2000
    sapmodel.file_Save(model_path)
    sapmodel.closeModel()

    # --- 4. 輸出週期結果 ---
    print("\n--- Calculated Periods ---")
    for group_name, data in period_x.items():
        print(f"Direction: X, Group: {group_name:<10} Period: {data['period']:.4f} s")
    for group_name, data in period_y.items():
        print(f"Direction: Y, Group: {group_name:<10} Period: {data['period']:.4f} s")
    for group_name, data in period_z.items():
        print(f"Direction: Z, Group: {group_name:<10} Period: {data['period']:.4f} s")

    output_excel_path = os.path.join(os.path.dirname(model_path), '01_period_results.xlsx')
    export_results_to_excel(period_x, period_y, period_z, output_excel_path)


if __name__ == "__main__":
    # TODO: 此處先指定GROUP，後續須配合UI傳入變數變換
    groups_to_run_x = ['ALL', 'Pier1']
    groups_to_run_y = ['ALL']
    groups_to_run_z = ['ALL']

    # 執行主流程
    run_analysis_and_export_results(
        MODEL_PATH, 
        groups_to_run_x, 
        groups_to_run_y, 
        groups_to_run_z
    )
