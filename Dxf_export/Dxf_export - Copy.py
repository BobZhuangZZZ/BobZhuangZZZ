import VBAPI
import os
import time

def FlatViewDefinintion(filename,rev):
    accuratelist=[]
    for i in range(Currentviews.Count):
        if Currentviews.Item(i).GetSheetNumber()==1:
            accuratelist.append(i)

    CurrentviewFront=''
    CurrentSheetinfo=CurrentDrw.GetSheetInfo(1)
    TranslateYs = {
        431.8: 98.345588235294,
        558.8: 35.5113636363634,
        863.6: 98.345588235294,
        1117.6: 35.5113636363634
    }
    TranslateY=TranslateYs[CurrentSheetinfo.Width]

    if CurrentDrw.ListModels().Count==1:
        if Currentviews.Count==1:
            CurrentviewFront=Currentviews.Item(0)
        else:
            i=0
            # 过滤出带flat的视图
            for i in accuratelist:
                try:
                    if ((Currentviews.Item(i).Outline.Item(0).Item(1) > TranslateY)
                            and
                        (abs(
                            Currentviews.Item(i).Outline.Item(1).Item(2) -
                            Currentviews.Item(i).Outline.Item(0).Item(2))/Currentviews.Item(i).Scale< 20)):
                        CurrentviewFront = Currentviews.Item(i)
                    else:
                        Currentviews.Item(i).Delete(True)
                except:
                    continue
    else:
        i=0
        filter_result_list=[]
        # 过滤出带flat的视图
        for i in accuratelist:
            try:
                if ((Currentviews.Item(i).Outline.Item(0).Item(1)>TranslateY)
                    and (Currentviews.Item(i).GetModel().GenericName)):
                    CurrentviewFront=Currentviews.Item(i)
                    filter_result_list.append(i)
                else:
                    Currentviews.Item(i).Delete(True)
            except:
                continue
        # 排除多个flat视图在图纸内部，根据视图面积最大值过滤出需要打印的展开图
        Targetindex=0
        MaxArea=0
        Maxwidth=0
        if len(filter_result_list)>1:
            for i in range(len(filter_result_list)):
                thisview=Currentviews.Item(filter_result_list[i])
                thisviewWidth=(thisview.Outline.Item(1).Item(0)-thisview.Outline.Item(0).Item(0))/thisview.Scale
                thisviewheight=(thisview.Outline.Item(1).Item(1)-thisview.Outline.Item(0).Item(1))/thisview.Scale
                # print(thisviewWidth*thisviewheight)
                # print(leftviewWidth*leftviewheight)
                if  (thisviewWidth>Maxwidth):
                    Maxwidth=thisviewWidth
                    Targetindex=i
                    MaxArea=thisviewWidth*thisviewheight
                elif(abs(thisviewWidth - Maxwidth)<1):
                    if (thisviewWidth*thisviewheight) > MaxArea:
                        Maxwidth=thisviewWidth
                        Targetindex=i
                        MaxArea=thisviewWidth*thisviewheight

            # print(Currentviews.Item(filter_result_list[Targetindex]).Name)
            for i in range(len(filter_result_list)):
                try:
                    if i==Targetindex:
                        CurrentviewFront=Currentviews.Item(filter_result_list[i])
                    else:
                        Currentviews.Item(filter_result_list[i]).Delete(True)
                except:
                    continue

    #根据flat视图当前的位置推算坐标原点
    if CurrentviewFront == '':
        CurrentDrw.Erase()
        # CurrentDrw.EraseWithDependencies()
        return CurrentviewFront
    else:
        Drawing_item_clean_up()
        CurrentviewFront.Name=filename
        X1=CurrentviewFront.Outline.Item(0).Item(0)
        Y1=CurrentviewFront.Outline.Item(0).Item(1)
        # Front_transform=VBAPI.CCpfcTransform3D().Create()
        # Front_X_vector = Front_transform.GetXAxis()
        # Front_X_vector.Set(0,-X1)
        # Front_X_vector.Set(1,-Y1+TranslateY)
        Translate_vector=VBAPI.CpfcVector3D()
        Translate_vector.Set(0,-X1)
        Translate_vector.Set(1,-Y1+TranslateY)
        #B	431.8	279.4	98.345588235294
        #C	558.8	431.8	35.5113636363634
        #D	863.6	558.8	98.345588235294
        #E	1117.6	863.6	35.5113636363634
        pdf_export_operation(filename,rev)
        CurrentviewFront.Translate(Translate_vector)
        # Drawing_item_clean_up()
        Dxf_export_operation(filename,rev)
        return CurrentviewFront

def FlatViewDefinintion_copperbar(filename,rev):
    Drawingpage=0
    targetindex=0
    CurrentviewFront=''
    # 查找展开图
    for i in range(Currentviews.Count):
        try:
            if ((Currentviews.Item(i).GetModel().GenericName) and
                (filename in Currentviews.Item(i).GetModel().InstanceName)
                    and (len(Currentviews.Item(i).Name) > 1)):
                Drawingpage=Currentviews.Item(i).GetSheetNumber()
                targetindex=i
                CurrentviewFront=Currentviews.Item(i)
                break
        except:
            continue

    # 定位到展开图的sheet
    CurrentDrw.CurrentSheetNumber=Drawingpage

    # 收集并删除当前sheet的剩余视图
    for i in range(Currentviews.Count):
        try:
            if ((Currentviews.Item(i).GetSheetNumber()==Drawingpage)and (i!=targetindex)):
                Currentviews.Item(i).Delete(True)
        except:
            continue

    CurrentSheetinfo=CurrentDrw.GetSheetInfo(1)
    TranslateYs = {
        431.8: 98.345588235294,
        558.8: 35.5113636363634,
        863.6: 98.345588235294,
        1117.6: 35.5113636363634
    }
    TranslateY=TranslateYs[CurrentSheetinfo.Width]

    #根据flat视图当前的位置推算坐标原点
    if CurrentviewFront == '':
        CurrentDrw.Erase()
        # CurrentDrw.EraseWithDependencies()
        return CurrentviewFront
    else:
        Drawing_item_clean_up()
        CurrentviewFront.Name=filename
        X1=CurrentviewFront.Outline.Item(0).Item(0)
        Y1=CurrentviewFront.Outline.Item(0).Item(1)
        # Front_transform=VBAPI.CCpfcTransform3D().Create()
        # Front_X_vector = Front_transform.GetXAxis()
        # Front_X_vector.Set(0,-X1)
        # Front_X_vector.Set(1,-Y1+TranslateY)
        Translate_vector=VBAPI.CpfcVector3D()
        Translate_vector.Set(0,-X1)
        Translate_vector.Set(1,-Y1+TranslateY)
        #B	431.8	279.4	98.345588235294
        #C	558.8	431.8	35.5113636363634
        #D	863.6	558.8	98.345588235294
        #E	1117.6	863.6	35.5113636363634
        change_view_display(CurrentviewFront)
        pdf_export_operation(filename,rev)
        CurrentviewFront.Translate(Translate_vector)
        # Drawing_item_clean_up()
        Dxf_export_operation(filename,rev)
        return CurrentviewFront

def Drawing_item_clean_up_old():
    session.RunMacro('''~ Command `ProCmdDwgPageSetup` ;
~ Activate `storage_conflicts` `OK_PushButton`;
~ Activate `pagesetup` `ChkShowFmt` 0;~ Activate `pagesetup` `OK`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Note;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdEditDeleteDwg`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Dimension`;
~ Activate `selspecdlg0` `SelScopeCheck` 1;
~ Select `selspecdlg0` `RuleTab` 1 `Attributes`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Drawing Item`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdDwgErase@PopupMenuGraphicWinStack` ;~ Activate `main_dlg_cur` `buffer_clean`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `RuleTab` 1 `Attributes`;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Draft Entity`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Drawing Table`;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdEditDeleteDwg`;
~ Activate `storage_conflicts` `OK_PushButton`;
~ Command `ProCmdDwgUpdateSheets`
''')
    CurrentDrw.Regenerate()
    time.sleep(3)

def Drawing_item_clean_up():
    session.RunMacro('''~ Command `ProCmdDwgPageSetup` ;
~ Activate `pagesetup` `ChkShowFmt` 0;~ Activate `pagesetup` `OK`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Drawing Table`;
~ Select `selspecdlg0` `RuleTab` 1 `Misc`;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdEditDeleteDwg`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Drawing Item`;
~ Select `selspecdlg0` `RuleTab` 1 `Misc`;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Draft Entity`;
~ Select `selspecdlg0` `RuleTab` 1 `Misc`;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdEditDeleteDwg`;
~ Command `ProCmdMdlTreeSearch` ;
~ Select `selspecdlg0` `SelOptionRadio` 1 `Dimension`;
~ Activate `selspecdlg0` `SelScopeCheck` 1;
~ Select `selspecdlg0` `RuleTab` 1 `Misc`;
~ Select `selspecdlg0` `RuleTypes` 1 `All`;
~ Activate `selspecdlg0` `EvaluateBtn`;
~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `UI Message Dialog` `ok`;
~ Activate `selspecdlg0` `CancelButton`;
~ Command `ProCmdEditDeleteDwg`
''')
    CurrentDrw.Regenerate()
    # time.sleep(3)

# ~ Command `ProCmdDwgUpdateSheets`
def Dxf_export_operation(Name,rev):
    DxfexoprtType=VBAPI.CCpfcDXFExportInstructions().Create()
    DxfName = str(Name) + '_'+ rev
    CurrentDrw.Export(Dxflink + DxfName, DxfexoprtType)
    Mywindow.Close()
    CurrentDrw.EraseWithDependencies()
    # session.RunMacro("~ Command `ProCmdModelErase` ;~ Activate `file_erase` `OK`")
    while True:
        if os.path.exists(Dxflink + str.lower(DxfName)+'.dxf'):
            os.rename(Dxflink + str.lower(DxfName)+'.dxf', Dxflink + str.upper(DxfName) +'.dxf')
            break

def pdf_export_operation(Name,rev):

    PDFexportoption=VBAPI.CCpfcPDFOption().Create()
    PDFexportoption.OptionType=VBAPI.CCpfcPDFOptionType().PDFOPT_SHEETS
    PDFexportoption.OptionValue=VBAPI.CMpfcArgument().CreateIntArgValue(0)

    PDFexportoption2=VBAPI.CCpfcPDFOption().Create()
    PDFexportoption2.OptionType=VBAPI.CCpfcPDFOptionType().PDFOPT_COLOR_DEPTH
    PDFexportoption2.OptionValue=VBAPI.CMpfcArgument().CreateIntArgValue(2)

    PDFexportoption3=VBAPI.CCpfcPDFOption().Create()
    PDFexportoption3.OptionType=VBAPI.CCpfcPDFOptionType().PDFOPT_LAUNCH_VIEWER
    PDFexportoption3.OptionValue=VBAPI.CMpfcArgument().CreateBoolArgValue(0)

    PDFexportoptions=VBAPI.CpfcPDFOptions()
    PDFexportoptions.Append(PDFexportoption)
    PDFexportoptions.Append(PDFexportoption2)
    PDFexportoptions.Append(PDFexportoption3)

    pdfexoprtType=VBAPI.CCpfcPDFExportInstructions().Create()
    pdfName = str(Name) + '_'+ rev
    pdfexoprtType.FilePath=pdfName
    pdfexoprtType.Options=PDFexportoptions
    pdfexoprtType.ProfilePath=None
    CurrentDrw.Export(Dxflink + pdfName, pdfexoprtType)
    os.rename(Dxflink + str.lower(pdfName)+'.pdf', Dxflink + pdfName +'.pdf')

def reset_print_area(CurrentviewFront):
    Thissheetinfo=CurrentDrw.GetSheetInfo(1)
    newareawidth=((CurrentviewFront.Outline.Item(1).Item(0)-CurrentviewFront.Outline.Item(0).Item(0))/(1000/Thissheetinfo.Width))+5
    newareaheight=((CurrentviewFront.Outline.Item(1).Item(1)-CurrentviewFront.Outline.Item(0).Item(1))/(1000/Thissheetinfo.Width))+5
    reset_print_area_Macro='''~ Command `ProCmdDwgPageSetup` ;
                              ~ Select `pagesetup` `TblFormats` 2 `0` `fmt`;
                              ~ Select `pagesetup` `TblFormats_INPUT` 1 `Variable...`;
                              ~ Update `pagesetup` `InpWidth` `''' + str(round(newareawidth,1)) + '''`;
                              ~ FocusOut `pagesetup` `InpWidth`;
                              ~ Update `pagesetup` `InpHeight` `''' + str(round(newareaheight,1)) + '''`;
                              ~ FocusOut `pagesetup` `InpHeight`;
                              ~ Activate `pagesetup` `OK`;
                              ~ Activate `keep_format_tables` `RemoveAll`;
                              ~ Activate `0_std_confirm` `OK`'''
    return reset_print_area_Macro

def change_view_display(view:VBAPI.IDpfcView2D):
    # 设置viewdisplay显示方式
    config_nohidden=VBAPI.CCpfcDisplayStyle().DISPSTYLE_NO_HIDDEN
    config_none=VBAPI.CCpfcTangentEdgeDisplayStyle().TANEDGE_NONE
    CableStyle=VBAPI.CCpfcCableDisplayStyle().CABLEDISP_DEFAULT
    # 创建viewdisplay并修改
    mynewstyle = VBAPI.CCpfcViewDisplay().Create(
    config_nohidden, config_none, CableStyle,
    True,False,True)
    view.Display=mynewstyle

def del_files(path):
    for root , dirs, files in os.walk(path):
        for name in files:
            if name.endswith(".log.1"):   #指定要删除的格式，这里是jpg 可以换成其他格式
                os.remove(os.path.join(root, name))
                # print ("Delete File: " + os.path.join(root, name))

def check_Save_Drawing():
    try:
        Drawingrevision=CurrentDrw.GetParam("PTC_MODIFIED")
        if Drawingrevision.Value.BoolValue==True:
            CurrentDrw.Save()
    except:
        CurrentDrw.Save()

def Export_the_dxf_for_Drawing(filename,rev):
    global CurrentDrw,Currentviews,CurrentviewFront,Mywindow
    # Open drawing
    descDrwOpen = VBAPI.CCpfcModelDescriptor()
    Drwname = str(filename)[0:8] + '.drw'
    descDrw= descDrwOpen.CreateFromFileName(Drwname)
    try:
        CurrentDrw=session.RetrieveModel(descDrw)
    except:
        CurrentviewFront=""
        return CurrentviewFront
    # CurrentDrw=session.GetModelFromDescr(descDrw)
    # check_Save_Drawing()
    Mywindow=session.OpenFile(descDrw)
    Mywindow.Activate()
    session.RunMacro('''~ Command `ProCmdDwgCrStdNewRefDim` ;
    ~ Activate `storage_conflicts` `OK_PushButton`;
    ~ Activate `dtl_attref_ui` `psh_cancel`
    ''')
    # session.RunMacro('''~ Command `ProCmdDwgCrStdNewRefDim` ;
    # ~ Activate `storage_conflicts` `OK_PushButton`;
    # ~ Activate `dtl_attref_ui` `psh_cancel`
    # ''')
    # if CurrentDrw.NumberOfSheets ==1:
    #     pass
    # else:
    #     CurrentDrw.CurrentSheetNumber=1
    #     for i in range(2,CurrentDrw.NumberOfSheets+1):
    #         CurrentDrw.DeleteSheet(2)
    Currentviews=CurrentDrw.List2DViews()
    # CurrentTransform=CurrentDrw.GetSheetTransform(1)
    # print(CurrentTransform.GetXAxis().Item(0),CurrentTransform.GetXAxis().Item(1),CurrentTransform.GetXAxis().Item(2))
    # print(CurrentTransform.GetYAxis().Item(0),CurrentTransform.GetYAxis().Item(1),CurrentTransform.GetYAxis().Item(2))
    # print(CurrentTransform.GetZAxis().Item(0),CurrentTransform.GetZAxis().Item(1),CurrentTransform.GetZAxis().Item(2))
    # print(Currentviews.Item(2).Outline.Item(1).Item(0)-Currentviews.Item(2).Outline.Item(0).Item(0),Currentviews.Item(2).Outline.Item(1).Item(1)-Currentviews.Item(2).Outline.Item(0).Item(1))
    # print(Currentviews.Item(2).Scale)
    # Mywindow.Refresh()
    CurrentviewFront=FlatViewDefinintion_copperbar(filename,rev)
    # Workspace_user.UndoCheckout(CurrentDrw)
    return CurrentviewFront
    # asyncConnection.Disconnect(2)

Dxflink=os.path.expanduser('~').replace('\\','/')+'/Downloads/dxf_export_Creo/'
Partlist=[]
RevList=[]
newlinelist=[]
for line in open(Dxflink + 'Part_list.txt',mode="r"):
    newlinelist=line.split(" ")
    Partlist.append(newlinelist[0])
    RevList.append(newlinelist[1].replace('\n',''))

# Partlist = [
#     571292150001
# ]
# RevList=['A']

fileopen = open(
    Dxflink + 'DXF_print_result.txt',
    mode="w",
    encoding="utf-8",
)

cAC = VBAPI.CCpfcAsyncConnection()
asyncConnection = cAC.Connect()
session = asyncConnection.Session
# session.ChangeDirectory('C:\\Users\\irdgff\\Downloads\\drw_sm\\')
session.SetConfigOption('dxf_out_drawing_scale','no')
session.SetConfigOption('dxf_out_scale_views','yes')
session.SetConfigOption('display_planes','no')
session.SetConfigOption('display_axes','no')
session.SetConfigOption('datum_point_display','no')
session.SetConfigOption('display_coord_sys','no')
session.SetConfigOption('drawing_warn_if_flex_feature','no')
#session.SetConfigOption('dm_checkout_on_the_fly','continue')
session.SetConfigOption('open_draw_simp_rep_by_default','no')
# try:
# Export_the_dxf_for_Drawing(12104527)
for Part,rev in zip(Partlist,RevList):
    Result = Export_the_dxf_for_Drawing(Part,rev)
    if Result == '':
        fileopen.writelines(str(Part)+' '+'dxf export failure\n')
    else:
        fileopen.writelines(str(Part)+' '+'dxf export success\n')
asyncConnection.Disconnect(2)
fileopen.close()
del_files(Dxflink)
os.remove(Dxflink + 'Part_list.txt')
# os.system('explorer '+Dxflink.replace("/","\\")[:-1])
# os.system('notepad ' + Dxflink + 'DXF_print_result.txt')
# except:
# asyncConnection.Disconnect(2)
# print(traceback.format_exc())
