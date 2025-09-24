# -*- coding: utf-8 -*-
import win32com.client as win32
import os
import pandas as pd
import math
import pythoncom
import sys

# 專案設定
# 這裡請替換為你要開啟的模型檔案路徑
MODEL_PATH = r"D:\Users\63427\Desktop\Code\EQ2SAP\example\model_test.sdb"


class Sap2000(object):
    """
    一個用於與 SAP2000 應用程式進行 COM 互動的包裝類別。
    提供了開啟、儲存、分析模型以及獲取結果等多種方法。
    """
    def __init__(self):
        self.SapObject = None
        self.SapModel = None

    def initializeNewModel(self, unitsTag=12):
        """
        確保 SAP2000 實例已啟動並初始化一個新模型。
        此方法會先嘗試附加到一個正在運行的 SAP2000 實例，如果失敗，則會啟動一個新的實例。
        """
        try:
            # 嘗試獲取一個正在運行的 SAP2000 實例
            self.SapObject = win32.GetActiveObject("SAP2000.SapObject")
            print("[訊息]：已成功附加到現有的 SAP2000 實例。")
        except pythoncom.com_error:
            # 如果沒有正在運行的實例，則創建一個新的
            print("[訊息]：未找到正在運行的 SAP2000 實例，正在啟動新實例...")
            try:
                self.SapObject = win32.Dispatch("SAP2000.SapObject")
                self.SapObject.ApplicationStart()
            except pythoncom.com_error:
                print("[錯誤]：無法啟動 SAP2000。請檢查 SAP2000 是否已正確安裝，並手動關閉所有背景中的 'SAP2000.exe' 程序後再試。")
                exit(1)

        self.SapModel = self.SapObject.SapModel
        # 清除當前模型並初始化一個新模型
        self.SapModel.InitializeNewModel(unitsTag)
        print("[訊息]：SAP2000 模型已初始化。")

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

    def getCoordSystem(self):
        """
        ---get the name of the present coordinate system---
        """
        currentCoordSysName = self.SapModel.GetPresentCoordSystem()

        return currentCoordSysName
    
    def getModelIsLocked(self):
        """
        The function returns True if the model is locked and False if it is unlocked.
        """
        IsLocked = self.SapModel.GetModelIsLocked()

        return IsLocked
    
    def setModelIsLocked(self, locker):
        """
        The item is True if the model is to be locked and False if it is to be unlocked.
        參數：
            locker (bool): 模型鎖定狀態。
        """
        IsLocked = self.SapModel.SetModelIsLocked(locker)

        return IsLocked

    def clearSelection(self):
        """
        This function deselects all objects in the model. 
        It returns zero if the selection status is successfully set, otherwise it returns nonzero.
        """
        self.SapModel.SelectObj.ClearSelection()

    def selectGroup(self, groupname):
        """
        This function selects or deselects all objects in the specified group.
        """
        self.SapModel.SelectObj.Group(groupname)

    def getSelected(self):
        """
        This function retrieves a list of selected objects.
        """
        ret = self.SapModel.SelectObj.GetSelected()

        return ret
    
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

    def define_LoadPatterns_Add(self,name,myType,SelfWTMultiplier=0,AddLoadCase=True):
        """
        新增一個載重模式（Load Pattern）。

        參數：
            name (str): 載重模式名稱。
            myType (int): 載重模式型態，對應 eLoadPatternType 列舉值。
                1=DEAD, 2=SUPERDEAD, 3=LIVE, 4=REDUCELIVE, 5=QUAKE, 6=WIND, 7=SNOW, 8=OTHER, 9=MOVE, 10=TEMPERATURE,
                11=ROOFLIVE, 12=NOTIONAL, 13=PATTERNLIVE, 14=WAVE, 15=BRAKING, 16=CENTRIFUGAL, 17=FRICTION, 18=ICE,
                19=WINDONLIVELOAD, 20=HORIZONTALEARTHPRESSURE, 21=VERTICALEARTHPRESSURE, 22=EARTHSURCHARGE, 23=DOWNDRAG,
                24=VEHICLECOLLISION, 25=VESSELCOLLISION, 26=TEMPERATUREGRADIENT, 27=SETTLEMENT, 28=SHRINKAGE, 29=CREEP,
                30=WATERLOADPRESSURE, 31=LIVELOADSURCHARGE, 32=LOCKEDINFORCES, 33=PEDESTRIANLL, 34=PRESTRESS,
                35=HYPERSTATIC, 36=BOUYANCY, 37=STREAMFLOW, 38=IMPACT, 39=CONSTRUCTION。
            SelfWTMultiplier (float): 自重乘數。
            AddLoadCase (bool): 是否自動新增對應的線性靜力載重案例。
        """
        self.SapModel.LoadPatterns.Add(name,myType,SelfWTMultiplier,AddLoadCase)


    def define_LoadCases_StaticLinear_SetCase(self, name):
        """
        初始化一個靜力線性載重案例。

        參數：
            name (str): 載重案例名稱（可為新建或既有）。
        """
        self.SapModel.LoadCases.StaticLinear.SetCase(name)

    def define_LoadCases_StaticLinear_SetLoads(self, name, numberLoads, loadType, loadName, scaleFactor):
        """
        設定指定分析案例的載重資料。

        參數：
            name (str): 既有靜力線性載重案例名稱。
            numberLoads (int): 指定案例的載重數量。
            loadType (list of str): 載重型態（'Load' 或 'Accel'）。
            loadName (list of str): 載重名稱。若 loadType 為 'Load'，此為已定義載重名稱；若為 'Accel'，此為 UX、UY、UZ、RX、RY 或 RZ。
            scaleFactor (list of float): 各載重的比例因子。對於 Accel UX/UY/UZ 單位為 L/s²，其餘無單位。
        """
        self.SapModel.LoadCases.StaticLinear.SetLoads(name, numberLoads, loadType, loadName, scaleFactor)

    def deltet_Pointobj_Deleteloadforce(self, name, loadpat, ItemTypeElm):
        """
        This function deletes all point load assignments, for the specified load pattern, from the specified point object(s).

        參數：
            Name (str): The name of a point object or a group depending on the value selected for ItemType item.
            LoadPat (str): The name of a defined load pattern.
            ItemType(int): This is one of the following items from the eItemType enumeration.
                            Object = 0
                            Group = 1
                            SelectedObjects = 2
        """
        self.SapModel.PointObj.DeleteLoadForce(name, loadpat, ItemTypeElm)

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

    def assign_PointObj_SetLoadForce(self,name,loadPat,value,Replace=False,CSys="Global",ItemType=0):
        """
        指定點物件的節點力。

        參數：
            name (str): 點物件名稱或群組名稱，依 ItemType 決定。
            loadPat (str): 載重模式名稱。
            value (list of float): 六個分量的點載重值。
                value[0]: F1 [力]
                value[1]: F2 [力]
                value[2]: F3 [力]
                value[3]: M1 [力×長度]
                value[4]: M2 [力×長度]
                value[5]: M3 [力×長度]
            Replace (bool): 若為 True，則先刪除舊有載重再指定新載重。
            CSys (str): 載重所用座標系統名稱，預設為 Global。
            ItemType (int): 指定對象型態：0=Object，1=Group，2=SelectedObjects。
                0: 指定 name 為單一點物件。
                1: 指定 name 為群組。
                2: 指定所有已選取點物件，忽略 name。
        """
        self.SapModel.PointObj.SetLoadForce(name,loadPat,value,Replace,CSys,ItemType)


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

def cal_eqforce(jointdisp, jointmass, group, eqfactor, vpa):
    dict_eqforce = {}
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
        all_mass = sum(mass_by_node.values())
        baseshear = all_mass * 9.81 * eqfactor[group.index(gp)]
        baseshear_vpa = baseshear * vpa

        # 計算節點地震力
        # [sum(wu)/sum(wuu)]*wu*(V/W) = (beta/zeta)*wu*eqfactor
        eqf_temp = {node: (beta/zeta) * eqfactor[group.index(gp)]*9.81 * mass_by_node[node] * disp_by_node[node] for node in disp_by_node.keys() & mass_by_node.keys()}

        sumeqf = abs(sum(eqf_temp.values()))
        if sumeqf >= baseshear_vpa:
            eqf = eqf_temp
        else:
            print("[警告]：第一振態分佈力總和未達總基底剪力{}%！".format(vpa*100))
            scaling_factor = baseshear_vpa / sumeqf
            eqf = {node: force * scaling_factor for node, force in eqf_temp.items()}

        dict_eqforce[gp] = {}
        dict_eqforce[gp]['beta'] = beta
        dict_eqforce[gp]['zeta'] = zeta
        dict_eqforce[gp]['eqfactor'] = eqfactor[group.index(gp)]
        dict_eqforce[gp]['TotalMass'] = all_mass
        dict_eqforce[gp]['BaseShear'] = baseshear
        dict_eqforce[gp]['mass'] = mass_by_node
        dict_eqforce[gp]['disp'] = disp_by_node
        dict_eqforce[gp]['wuu'] = wuu
        dict_eqforce[gp]['wu'] = wu
        dict_eqforce[gp]['eqforce_origin'] = eqf_temp
        dict_eqforce[gp]['eqforce'] = eqf

    return dict_eqforce

def cal_eqvforce(jointdisp, jointmass, group, eqfactor):
    dict_eqforce = {}
    # 默認第一組為上構，第二組為下構
    for gp in group:
        dict_mass = jointmass[gp]
        mass_by_node = dict(zip(dict_mass['node_name'], dict_mass['node_mass']))
        # 計算節點地震力
        eqf = {node: eqfactor[group.index(gp)]*9.81 * mass_by_node[node] for node in mass_by_node.keys()}
        
        dict_eqforce[gp] = {}
        dict_eqforce[gp]['eqfactor'] = eqfactor[group.index(gp)]
        dict_eqforce[gp]['mass'] = mass_by_node
        dict_eqforce[gp]['eqforce'] = eqf

    return dict_eqforce

def merge_group_data(data_dict, new_group_name):
    """
    將來自多個群組的資料字典合併為單一群組。
    處理重複節點時，會保留第一個遇到的節點資料。

    Args:
        data_dict (dict): 來自 get_disp 或 get_mass 的字典，例如 {'Pier1': {...}, 'Pier2': {...}}。
        new_group_name (str): 新合併群組的名稱。
        
    Returns:
        dict: 包含單一合併群組條目的字典，例如 {'Piers_Combined': {...}}。
    """
    if not data_dict:
        return {new_group_name: {'node_num': 0, 'node_name': [], 'node_data': []}}

    # 確定資料是位移('node_disp')還是質量('node_mass')
    first_group_data = next(iter(data_dict.values()))
    data_key = 'node_disp' if 'node_disp' in first_group_data else 'node_mass'

    all_node_names = []
    all_node_data = []

    for group_data in data_dict.values():
        all_node_names.extend(group_data.get('node_name', []))
        all_node_data.extend(group_data.get(data_key, []))

    # 透過字典來處理重複節點，保留第一個出現的值
    unique_nodes = {}
    for name, data_val in zip(all_node_names, all_node_data):
        if name not in unique_nodes:
            unique_nodes[name] = data_val
            
    merged_data = {
        'node_num': len(unique_nodes),
        'node_name': list(unique_nodes.keys()),
        data_key: list(unique_nodes.values())
    }
    
    return {new_group_name: merged_data}

def merge_force_data(force_data_dict):
    """
    將來自多個群組的 'eqforce' 字典合併為單一字典。

    Args:
        force_data_dict (dict): 一個字典，其鍵為群組名稱，值為包含 'eqforce' 字典的字典。
                                e.g., {'Group1': {'eqforce': {'Node1': 10}}, 'Group2': {'eqforce': {'Node2': 20}}}

    Returns:
        dict: 一個包含所有合併後節點力的單一字典。
              e.g., {'Node1': 10, 'Node2': 20}
    """
    merged_forces = {}
    for group_data in force_data_dict.values():
        if 'eqforce' in group_data:
            merged_forces.update(group_data['eqforce'])
    return merged_forces

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

def export_eqforce_to_excel(eqforce_x, eqforce_y, eqforce_z, output_path):
    """
    將計算出的地震力結果彙整並輸出至包含多個工作表的 Excel 檔案。

    - 'EQForce Summary': 各群組與方向的總地震力。
    - 'Group-Direction': 各群組與方向的詳細節點力、質量與位移。

    參數:
        eqforce_x (dict): X 方向的地震力計算結果。
        eqforce_y (dict): Y 方向的地震力計算結果。
        eqforce_z (dict): Z 方向的地震力計算結果。
        output_path (str): Excel 檔案的完整輸出路徑。
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 1. 彙整總覽結果
        summary_data = []
        for direction, force_data in [('X', eqforce_x), ('Y', eqforce_y), ('Z', eqforce_z)]:
            for group_name, data in force_data.items():
                total_force = sum(data.get('eqforce', {}).values())
                summary_data.append({
                    'Group': group_name,
                    'Direction': direction,
                    'Total Force': total_force,
                    'EQ Factor': data.get('eqfactor'),
                    'Sum(wu) (beta)': data.get('beta'),
                    'Sum(wuu) (zeta)': data.get('zeta'),
                    'Total Mass': data.get('TotalMass'),
                    'Base Shear': data.get('BaseShear')
                })
        df_summary = pd.DataFrame(summary_data)
        # 調整總覽表的欄位順序
        summary_cols = ['Group', 'Direction', 'Total Force', 'EQ Factor', 'Sum(wu) (beta)', 'Sum(wuu) (zeta)', 'Total Mass', 'Base Shear']
        # 篩選掉值全為 None 的欄位 (例如 Z 方向沒有 beta, zeta)
        df_summary = df_summary.dropna(axis=1, how='all')
        # 確保欄位順序正確
        ordered_cols = [col for col in summary_cols if col in df_summary.columns]
        df_summary[ordered_cols].to_excel(writer, sheet_name='EQForce Summary', index=False, float_format="%.4f")

        # 2. 遍歷每個方向和群組，寫入詳細的節點力工作表
        for direction, force_data in [('X', eqforce_x), ('Y', eqforce_y), ('Z', eqforce_z)]:
            for group_name, data in force_data.items():
                node_forces = data.get('eqforce', {})
                node_forces_origin = data.get('eqforce_origin', {})
                node_masses = data.get('mass', {})
                node_disps = data.get('disp', {})  # Z 方向的 data 中沒有 'disp'
                node_wu = data.get('wu', {})      # X, Y 方向的中間計算值
                node_wuu = data.get('wuu', {})    # X, Y 方向的中間計算值

                all_nodes = sorted(
                    list(node_forces.keys()),
                    key=lambda x: (0, int(x)) if x.isdigit() else (1, x)
                )

                detail_data = []
                for node in all_nodes:
                    row = {
                        'Node': node,
                        'Mass': node_masses.get(node),
                        'Force': node_forces.get(node)
                    }
                    # 只有 X, Y 方向的計算結果包含位移
                    if node_disps:
                        row['Displacement'] = node_disps.get(node)
                    if node_wu:
                        row['wu'] = node_wu.get(node)
                    if node_wuu:
                        row['wuu'] = node_wuu.get(node)
                    if node_forces_origin:
                        row['Force_Origin'] = node_forces_origin.get(node)
                    detail_data.append(row)

                if not detail_data:
                    continue

                df_details = pd.DataFrame(detail_data)
                # 調整欄位順序，確保 Displacement 在 Force 之前
                detail_cols = ['Node', 'Mass']
                if 'Displacement' in df_details.columns:
                    detail_cols.append('Displacement')
                if 'wu' in df_details.columns:
                    detail_cols.append('wu')
                if 'wuu' in df_details.columns:
                    detail_cols.append('wuu')
                if 'Force_Origin' in df_details.columns:
                    detail_cols.append('Force_Origin')
                detail_cols.append('Force')
                df_details = df_details[detail_cols]

                sheet_name = f"{group_name}-{direction}"[:31]
                df_details.to_excel(writer, index=False, sheet_name=sheet_name, float_format="%.6e")

    print(f"[訊息]：地震力計算結果已成功匯出至：{output_path}")

def setup_and_run_sap_analysis(model_path):
    """
    開啟 SAP2000 模型，設定並執行分析。

    1. 開啟並準備 SAP2000 模型。
    2. 設定並執行單位力載重分析。

    參數:
        model_path (str): SAP2000 模型檔案的完整路徑。

    回傳:
        Sap2000: 已執行分析的 Sap2000 物件。
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
    status_lock = sapmodel.getModelIsLocked()
    if status_lock:
        sapmodel.setModelIsLocked(False)
    lc_unit = {"UNIT-X": "UX", "UNIT-Y": "UY", "UNIT-Z": "UZ"}
    for lc, dir_name in lc_unit.items():
        sapmodel.define_LoadCases_StaticLinear_SetCase(lc)
        sapmodel.define_LoadCases_StaticLinear_SetLoads(lc, 1, ["Accel"], [dir_name], [9.81])
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
    
    return sapmodel

def run_analysis_period(model_path, groups_x, groups_y, groups_z):
    """
    執行完整的 SAP2000 週期分析與結果匯出流程。

    1. 開啟模型並執行分析。
    2. 提取位移與質量，計算各群組與方向的週期。
    3. 將結果印至主控台並匯出至 Excel 檔案。

    參數:
        model_path (str): SAP2000 模型檔案的完整路徑。
        groups_x (list): 要在 X 方向分析的群組名稱列表。
        groups_y (list): 要在 Y 方向分析的群組名稱列表。
        groups_z (list): 要在 Z 方向分析的群組名稱列表。
    """
    # --- 1. 開啟模型並執行分析 ---
    sapmodel = setup_and_run_sap_analysis(model_path)

    # --- 2. 獲取分析結果並計算週期 ---
    # 獲取 X 和 Y 方向的位移與質量
    jointdisp_x = get_disp(sapmodel, 'UNIT-X', groups_x, 7)
    jointdisp_y = get_disp(sapmodel, 'UNIT-Y', groups_y, 8)
    jointmass_x = get_mass(sapmodel, groups_x, 3)
    jointmass_y = get_mass(sapmodel, groups_y, 4)

    # 對於 Z 方向，獲取結果後合併為一個群組
    raw_jointdisp_z = get_disp(sapmodel, 'UNIT-Z', groups_z, 9)
    raw_jointmass_z = get_mass(sapmodel, groups_z, 5)
    merged_z_group_name = 'StructZdir'
    jointdisp_z = merge_group_data(raw_jointdisp_z, merged_z_group_name)
    jointmass_z = merge_group_data(raw_jointmass_z, merged_z_group_name)

    period_x = cal_period(jointdisp_x, jointmass_x, groups_x)
    period_y = cal_period(jointdisp_y, jointmass_y, groups_y)
    # Z 方向週期計算使用合併後的單一群組
    period_z = cal_period(jointdisp_z, jointmass_z, [merged_z_group_name])
    
    print("[訊息]：週期計算完成！")

    # 關閉 SAP2000
    sapmodel.file_Save(model_path)
    sapmodel.closeModel()

    # --- 3. 輸出週期結果 ---
    print("\n--- Calculated Periods ---")
    for group_name, data in period_x.items():
        print(f"Direction: X, Group: {group_name:<10} Period: {data['period']:.4f} s")
    for group_name, data in period_y.items():
        print(f"Direction: Y, Group: {group_name:<10} Period: {data['period']:.4f} s")
    for group_name, data in period_z.items():
        print(f"Direction: Z, Group: {group_name:<10} Period: {data['period']:.4f} s")

    output_excel_path = os.path.join(os.path.dirname(model_path), '01_period_results.xlsx')
    export_results_to_excel(period_x, period_y, period_z, output_excel_path)
    print("[訊息]：計算輸出完成！")

def run_analysis_eqforce(model_path, groups_x, groups_y, groups_z, eqfactor_x, eqfactor_y, eqfactor_z, v_percent):
    # --- 1. 開啟模型並執行分析 ---
    sapmodel = setup_and_run_sap_analysis(model_path)

    # --- 2. 獲取分析結果 ---
    # 獲取 X 和 Y 方向的位移與質量
    jointdisp_x = get_disp(sapmodel, 'UNIT-X', groups_x, 7)
    jointdisp_y = get_disp(sapmodel, 'UNIT-Y', groups_y, 8)
    jointdisp_z = get_disp(sapmodel, 'UNIT-Z', groups_z, 9)
    jointmass_x = get_mass(sapmodel, groups_x, 3)
    jointmass_y = get_mass(sapmodel, groups_y, 4)
    jointmass_z = get_mass(sapmodel, groups_z, 5)

    # --- 3. 計算分析橫力 ---
    # 計算X, Y方向地震節點力
    eqforce_x = cal_eqforce(jointdisp_x, jointmass_x, groups_x, eqfactor_x, v_percent)
    eqforce_y = cal_eqforce(jointdisp_y, jointmass_y, groups_y, eqfactor_y, v_percent)

    # 將各群組的地震力合併為單一字典
    EQF_x = merge_force_data(eqforce_x)
    EQF_y = merge_force_data(eqforce_y)
    
    # --- 4. 計算分析垂直力 ---
    # 計算Z方向地震節點力
    eqforce_z = cal_eqvforce(jointdisp_z, jointmass_z, groups_z, eqfactor_z)
    
    # 將各群組的地震力合併為單一字典
    EQF_z = merge_force_data(eqforce_z)
    print("[訊息]：分布力計算完成！")

    # --- 5. Assign地震力 ---
    presentcoordsystem = sapmodel.getCoordSystem()
    status_lock = sapmodel.getModelIsLocked()
    if status_lock:
        sapmodel.setModelIsLocked(False)
    # EQ Load Cases設定
    lc_list = sapmodel.loadcases_getnamelist()
    for case in ['EQL', 'EQT', 'EQV']:
        if case not in lc_list[2]:
            sapmodel.define_LoadPatterns_Add(case,5)
            sapmodel.define_LoadCases_StaticLinear_SetCase(case)
            sapmodel.define_LoadCases_StaticLinear_SetLoads(case, 1, ["Load"], [case], [1])

    # 地震力加載
    def eqforce_apply(lclabel, group, EQF, presentcoordsystem):
        sapmodel.deltet_Pointobj_Deleteloadforce('ALL', lclabel, 1)
        sapmodel.clearSelection()
        for sg in group:
            sapmodel.selectGroup(sg)
        res = sapmodel.getSelected()
        objdict = dict(zip(res[3], res[2])) # res[3]為object name, res[2]為object type
        for objname, objtype in objdict.items():
            if objtype == 1:
                value = EQF.get(objname)
                if value is None:
                    continue
                if lclabel == 'EQL':
                    forcedof = [value,0,0,0,0,0]
                elif lclabel == 'EQT':
                    forcedof = [0,value,0,0,0,0]
                elif lclabel == 'EQV':
                    forcedof = [0,0,value,0,0,0]
                else:
                    print('[錯誤]: 地震LoadCase命名有誤')
                    os._exit(-1)
                sapmodel.assign_PointObj_SetLoadForce(objname,lclabel,forcedof,Replace=True,CSys=presentcoordsystem,ItemType=0)

    eqforce_apply('EQL', groups_x, EQF_x, presentcoordsystem)
    eqforce_apply('EQT', groups_y, EQF_y, presentcoordsystem)
    eqforce_apply('EQV', groups_z, EQF_z, presentcoordsystem)
    print("[訊息]：地震力施加完成！")

    # 關閉 SAP2000
    sapmodel.file_Save(model_path)
    sapmodel.closeModel()
    print("[訊息]：SAP2000關閉。")

    # --- 6. 輸出地震力結果 ---
    output_excel_path = os.path.join(os.path.dirname(model_path), '02_eqforce_results.xlsx')
    export_eqforce_to_excel(eqforce_x, eqforce_y, eqforce_z, output_excel_path)

class Logger(object):
    """
    一個日誌記錄器類別，可以將輸出同時寫入終端和檔案。
    """
    def __init__(self, filename="pyeqlog.log", stream=sys.stdout):
        self.terminal = stream
        # 使用 utf-8 編碼以支援中文字元
        self.log = open(filename, "w", encoding='utf-8')

    def write(self, message):
        """將訊息寫入終端和日誌檔案。"""
        self.terminal.write(message)
        self.log.write(message)
        self.flush() # 確保即時寫入

    def flush(self):
        """刷新緩衝區，確保所有內容都已寫入。"""
        self.terminal.flush()
        self.log.flush()

    def __getattr__(self, attr):
        """將其他屬性請求代理給原始流對象。"""
        return getattr(self.terminal, attr)


if __name__ == "__main__":
    # TODO: 此處先指定GROUP，後續須配合UI傳入變數變換
    groups_to_run_x = ['Pier1','Pier2']
    groups_to_run_y = ['ALL']
    groups_to_run_z = ['ALL']

    eqfactor_to_run_x = [0.15, 0.15]
    eqfactor_to_run_y = [0.15]
    eqfactor_to_run_z = [0.15, 0.05]
    eqpercent = 0.9

    # --- 設定日誌記錄 ---
    # 將日誌檔案儲存在模型檔案所在的目錄下
    log_file_path = os.path.join(os.path.dirname(MODEL_PATH), 'pyeqlog.log')
    original_stdout = sys.stdout
    sys.stdout = Logger(log_file_path, original_stdout)

    try:
        # --- 執行主流程 ---
        # run_analysis_period(
        #     MODEL_PATH, 
        #     groups_to_run_x, 
        #     groups_to_run_y, 
        #     groups_to_run_z
        # )

        run_analysis_eqforce(
            MODEL_PATH, 
            groups_to_run_x, 
            groups_to_run_y, 
            groups_to_run_z,
            eqfactor_to_run_x, 
            eqfactor_to_run_y, 
            eqfactor_to_run_z,
            eqpercent
        )
    finally:
        # --- 還原標準輸出並關閉檔案 ---
        sys.stdout = original_stdout
        print(f"[訊息]：日誌已寫入 {log_file_path}")