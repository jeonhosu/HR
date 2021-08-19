using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace HRMF0518
{
    public partial class HRMF0518 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private string mCompany = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public HRMF0518()
        {
            InitializeComponent();
        }

        public HRMF0518(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_0.BringToFront();
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
             
            if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
            {
                string vPERSON_NUM = string.Empty;
                if (IGR_PAYMENT_SPREAD_DTL.RowIndex < 0)
                {
                    vPERSON_NUM = string.Empty;
                }
                else
                {
                    vPERSON_NUM = iString.ISNull(IGR_PAYMENT_SPREAD_DTL.GetCellValue("PERSON_NUM"));
                }

                IDA_PAYMENT_SPREAD_DTL.Fill();

                if (vPERSON_NUM == string.Empty)
                {
                    IGR_PAYMENT_SPREAD_DTL.Focus();
                }
                else
                {
                    int vIDX_PERSON_NUM = IGR_PAYMENT_SPREAD_DTL.GetColumnToIndex("PERSON_NUM");
                    for (int r = 0; r < IGR_PAYMENT_SPREAD_DTL.RowCount; r++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PAYMENT_SPREAD_DTL.GetCellValue(r, vIDX_PERSON_NUM)))
                        {
                            IGR_PAYMENT_SPREAD_DTL.CurrentCellMoveTo(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SPREAD_DTL.CurrentCellActivate(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SPREAD_DTL.Focus();
                            return;
                        }
                    }
                }
            }             
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM.TabIndex)
            {
                string vPERSON_NUM = string.Empty;
                if (IGR_PAYMENT_SUM_SHT.RowIndex < 0)
                {
                    vPERSON_NUM = string.Empty;
                }
                else
                {
                    vPERSON_NUM = iString.ISNull(IGR_PAYMENT_SUM_SHT.GetCellValue("PERSON_NUM"));
                }
                IDA_PAYMENT_SUM_SHT.Fill();

                if (vPERSON_NUM == string.Empty)
                {
                    IGR_PAYMENT_SUM_SHT.Focus();
                }
                else
                {
                    int vIDX_PERSON_NUM = IGR_PAYMENT_SUM_SHT.GetColumnToIndex("PERSON_NUM");
                    for (int r = 0; r < IGR_PAYMENT_SUM_SHT.RowCount; r++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PAYMENT_SUM_SHT.GetCellValue(r, vIDX_PERSON_NUM)))
                        {
                            IGR_PAYMENT_SUM_SHT.CurrentCellMoveTo(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SUM_SHT.CurrentCellActivate(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SUM_SHT.Focus();
                            return;
                        }
                    }
                }
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_II.TabIndex)
            {
                string vPERSON_NUM = string.Empty;
                if (IGR_PAYMENT_SUM_VCC.RowIndex < 0)
                {
                    vPERSON_NUM = string.Empty;
                }
                else
                {
                    vPERSON_NUM = iString.ISNull(IGR_PAYMENT_SUM_VCC.GetCellValue("PERSON_NUM"));
                }
                IDA_PAYMENT_SUM_VCC.Fill();

                if (vPERSON_NUM == string.Empty)
                {
                    IGR_PAYMENT_SUM_VCC.Focus();
                }
                else
                {
                    int vIDX_PERSON_NUM = IGR_PAYMENT_SUM_VCC.GetColumnToIndex("PERSON_NUM");
                    for (int r = 0; r < IGR_PAYMENT_SUM_VCC.RowCount; r++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PAYMENT_SUM_VCC.GetCellValue(r, vIDX_PERSON_NUM)))
                        {
                            IGR_PAYMENT_SUM_VCC.CurrentCellMoveTo(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SUM_VCC.CurrentCellActivate(r, vIDX_PERSON_NUM);
                            IGR_PAYMENT_SUM_VCC.Focus();
                            return;
                        }
                    }
                }
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_STATUS.TabIndex)
            {
                IDA_PAYMENT_SUM_STATUS.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_DEPT.TabIndex)
            {
                IDA_PAYMENT_SUM_DEPT.Fill();
            }
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        //private void XLPrinting_1(string pOutChoice, ISDataAdapter pAdapter)
        //{// pOutChoice : 출력구분.
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;

        //    object vToday = DateTime.Today.ToShortDateString();

        //    Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    Application.DoEvents();

        //    //출력구분이 파일인 경우 처리.
        //    if (pOutChoice == "FILE")
        //    {
        //        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //        vSaveFileName = string.Format("Accounts_{0}", vToday);

        //        saveFileDialog1.Title = "Excel Save";
        //        saveFileDialog1.FileName = vSaveFileName;
        //        saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //        saveFileDialog1.DefaultExt = "xls";
        //        if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //        {
        //            return;
        //        }
        //        else
        //        {
        //            vSaveFileName = saveFileDialog1.FileName;
        //            System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //            try
        //            {
        //                if (vFileName.Exists)
        //                {
        //                    vFileName.Delete();
        //                }
        //            }
        //            catch (Exception EX)
        //            {
        //                MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                return;
        //            }
        //        }
        //        vMessageText = string.Format(" Writing Starting...");
        //    }
        //    else
        //    {
        //        vMessageText = string.Format(" Printing Starting...");
        //    }
        //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //    Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    Application.DoEvents();

        //    int vPageNumber = 0;
        //    //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);
        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "HRMF0516_001.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            // 헤더 부분 인쇄.
        //            //xlPrinting.HeaderWrite(vAccountBook, vToday);

        //            // 라인 인쇄
        //            vPageNumber = xlPrinting.LineWrite(IGR_PAYMENT_ITEM_SUM);

        //            //출력구분에 따른 선택(인쇄 or file 저장)
        //            if (pOutChoice == "PRINT")
        //            {
        //                xlPrinting.Printing(1, vPageNumber);
        //            }
        //            else if (pOutChoice == "FILE")
        //            {
        //                xlPrinting.SAVE(vSaveFileName);
        //            }

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------

        //            vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        else
        //        {
        //            vMessageText = "Excel File Open Error";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        //-------------------------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        xlPrinting.Dispose();

        //        vMessageText = ex.Message;
        //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

        #endregion;
        
        #region ----- XL Print 1 Methods ----

        private void XLPrinting_Main()
        {
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(PAY_YYYYMM_0.EditValue));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0518");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));

            //print type 설정
            DialogResult vdlgResult;
            HRMF0518_PRINT_TYPE vHRMF0518_PRINT_TYPE = new HRMF0518_PRINT_TYPE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0518_PRINT_TYPE.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            string vPRINT_TYPE = iString.ISNull(vHRMF0518_PRINT_TYPE.Get_Printer_Type);
            if (vPRINT_TYPE == string.Empty)
            {
                return;
            }
            vHRMF0518_PRINT_TYPE.Dispose();

            //급상여대장인쇄.
            XLPrinting(vPRINT_TYPE);

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }
         
        private void XLPrinting(string pOutput_Type)
        {  
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
            {
                int vCountRowGrid = IGR_PAYMENT_SPREAD_DTL.RowCount;

                if (vCountRowGrid < 1)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                XLPrinting_DTL(pOutput_Type);
            }
               
            System.Windows.Forms.Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }


        //급상여 상세//
        private void XLPrinting_DTL(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vTitle = string.Empty;
            string vSaveFileName = string.Empty;

            int vPageNumber = 0;

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            if (pOutput_Type == "FILE")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "excel file(*.xls)|*.xls|(*.xlsx)|*.xlsx";
                vSaveFileDialog.DefaultExt = "xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
            }
            else if (pOutput_Type == "PDF")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "pdf file(*.pdf)|*.pdf";
                vSaveFileDialog.DefaultExt = "pdf";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
            }

            vMessageText = string.Format(" Printing Starting...");

            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents(); 

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0518_011.xlsx";
                //-------------------------------------------------------------------------------------

                bool IsOpen = xlPrinting.XLFileOpen();
                if (IsOpen == true)
                {
                    isAppInterfaceAdv1.OnAppMessage("Printing Start...");

                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    int vTerritory = GetTerritory(IGR_PAYMENT_SPREAD_DTL.TerritoryLanguage);

                    string vUserName = isAppInterfaceAdv1.AppInterface.LoginDescription;

                    string vCORP_NAME = CORP_NAME_0.EditValue as string;
                    string vYYYYMM = PAY_YYYYMM_0.EditValue as string;
                    string vWageTypeName = WAGE_TYPE_NAME_0.EditValue as string;
                    string vDepartment_NAME = DEPT_NAME_0.EditValue as string;

                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    vPageNumber = xlPrinting.XLWirteMain(IGR_PAYMENT_SPREAD_DTL, vLOCAL_DATE, vUserName, vCORP_NAME, vYYYYMM, vWageTypeName, vDepartment_NAME);
                    
                    if (pOutput_Type == "PDF")
                    { 
                        xlPrinting.PDF_Save(vSaveFileName);
                    }
                    else
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                }
                else
                {
                    xlPrinting.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }

            xlPrinting.Dispose();

            vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2007;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2007;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //ExportXL(igrMONTH_PAYMENT);
                    //XLPrinting("FILE");
                    if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_DETAIL.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_SPREAD_DTL);
                    }                     
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_SUM_SHT);
                    }
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_II.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_SUM_VCC);
                    }
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_STATUS.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_SUM_STATUS);
                    }
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SUM_DEPT.TabIndex)
                    {
                        ExcelExport(IGR_PAYMENT_SUM_DEPT);
                    }
                }
            }
        }

        #endregion;
        
        #region ----- Form Event -----

        private void HRMF0518_Load(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);

            DefaultCorporation();              //Default Corp. 
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaPAY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("PAY_TYPE", "Y");
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }
        
        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        #endregion


    }
}