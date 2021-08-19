using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0515
{
    public partial class HRMF0515 : Office2007Form
    {

        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;
        
        #region ----- Constructor -----

        public HRMF0515(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
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

        private void DefaultPaymentTerm()
        {
            // 조회년월 SETTING
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");

            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            idcYYYYMM_TERM.SetCommandParamValue("W_YYYYMM", PAY_YYYYMM_0.EditValue);
            idcYYYYMM_TERM.ExecuteNonQuery();
            START_DATE_0.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_START_DATE");
            END_DATE_0.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_END_DATE");
        }

        private bool AddAllowance_Check()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Year Month(급여 년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
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
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }

            if (iString.ISNull(WAGE_TYPE_0.EditValue) == String.Empty)
            {// 급상여구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            if (itbPAYMENT_ADDITION.SelectedTab.TabIndex == 1)
            {
                INIT_ALLOWANCE_COLUMN();
                idaADD_ALLOWANCE.Fill();
                igrADD_ALLOWANCE.Focus();
            }
            else if (itbPAYMENT_ADDITION.SelectedTab.TabIndex == 2)
            {
                INIT_DEDUCTION_COLUMN();
                idaADD_DEDUCTION.Fill();
                igrADD_DEDUCTION.Focus();
            }            
        }

        private void INIT_ALLOWANCE_COLUMN()
        {
            idaPROMPT_ALLOWANCE.Fill();
            if (idaPROMPT_ALLOWANCE.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            int mGRID_START_COL = 11;   // 그리드 시작 COLUMN.
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 39;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = idaPROMPT_ALLOWANCE.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = idaPROMPT_ALLOWANCE.CurrentRow[mIDX_Column];
                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    igrADD_ALLOWANCE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                }
                else
                {
                    igrADD_ALLOWANCE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    igrADD_ALLOWANCE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    igrADD_ALLOWANCE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }
            igrADD_ALLOWANCE.ResetDraw = true;
        }

        private void INIT_DEDUCTION_COLUMN()
        {
            idaPROMPT_DEDUCTION.Fill();
            if (idaPROMPT_DEDUCTION.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            int mGRID_START_COL = 11;   // 그리드 시작 COLUMN.
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 29;       // 종료 COLUMN.
            int mENABLED_COLUMN;        // 사용여부 COLUMN.

            object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                mENABLED_COLUMN = mMax_Column + mIDX_Column;
                mENABLED_FLAG = idaPROMPT_DEDUCTION.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = idaPROMPT_DEDUCTION.CurrentRow[mIDX_Column];
                if (iString.ISNull(mENABLED_FLAG, "N") == "N".ToString())
                {
                    igrADD_DEDUCTION.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                }
                else
                {
                    igrADD_DEDUCTION.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    igrADD_DEDUCTION.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    igrADD_DEDUCTION.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }
            igrADD_DEDUCTION.ResetDraw = true;
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        

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
                    if (idaADD_ALLOWANCE.IsFocused)
                    {
                        idaADD_ALLOWANCE.Cancel();
                    }
                    else if (idaADD_DEDUCTION.IsFocused)
                    {
                        idaADD_DEDUCTION.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaADD_ALLOWANCE.IsFocused)
                    {
                        idaADD_ALLOWANCE.Delete();
                    }
                    else if (idaADD_DEDUCTION.IsFocused)
                    {
                        idaADD_DEDUCTION.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0515_Load(object sender, EventArgs e)
        {
            idaADD_ALLOWANCE.FillSchema();
            idaADD_DEDUCTION.FillSchema();            
        }

        private void HRMF0515_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();       //Default Corp.
            DefaultPaymentTerm();       //Default Term.     
        }

        #endregion  

        #region ----- Adapter Event -----
        // Allowance 항목.
        private void idaADD_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Start Year Month(시작년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaADD_ALLOWANCE_PreDelete(ISPreDeleteEventArgs e)
        {
        }   
        #endregion

        #region ----- LookUp Event -----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        //#region ----- Excel UpLoad -----

        //private void bXL_Choice_Allowance_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    OpenXL();
        //}

        //private void bXL_UpLoad_Allowance_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    LoadingSTART();
        //}

        //private void bXL_Choice_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    OpenXL();
        //}

        //private void bXL_UpLoad_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        //{
        //    LoadingSTART();
        //}

        //#endregion

        //#region ----- Excel Open Method ----

        //private void OpenXL()
        //{
        //    string vMessage = string.Empty;

        //    try
        //    {
        //        System.IO.DirectoryInfo vOpenFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));

        //        openFileDialog1.RestoreDirectory = true;
        //        openFileDialog1.Title = "Excel_Open";
        //        openFileDialog1.DefaultExt = "xls";

        //        openFileDialog1.InitialDirectory = vOpenFolder.FullName;
        //        openFileDialog1.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files(*.*)|*.*";
        //        openFileDialog1.FileName = "*.xls;*.xlsx";
        //        if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //        {
        //            int vIndexTab = itbPAYMENT_ADDITION.SelectedIndex;
        //            if (vIndexTab == 0) //지급
        //            {
        //                ePath_Allowance.EditValue = openFileDialog1.FileName;
        //            }
        //            else if (vIndexTab == 1) //공제
        //            {
        //                ePath_Deduction.EditValue = openFileDialog1.FileName;
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //        System.Windows.Forms.Application.DoEvents();
        //    }
        //}

        //#endregion;

        //#region ----- Loading Start Method ----

        //private void LoadingSTART()
        //{
        //    string vMessage = string.Empty;
        //    bool isLoadXL_OK = false;

        //    string vOpenExcelFileName = openFileDialog1.FileName;

        //    bool isNull = string.IsNullOrEmpty(vOpenExcelFileName);
        //    if (isNull == true || vOpenExcelFileName == "*.xls;*.xlsx")
        //    {
        //        vMessage = string.Format("Excel file not selected");
        //        isAppInterfaceAdv1.OnAppMessage(vMessage);
        //        MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        System.Windows.Forms.Application.DoEvents();

        //        return;
        //    }

        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    try
        //    {
        //        int vIndexTab = itbPAYMENT_ADDITION.SelectedIndex;
        //        if (vIndexTab == 0) //지급
        //        {
        //            idaXLUPLOAD_ALLOWANCE.Cancel();

        //            ePath_Allowance.EditValue = vOpenExcelFileName;
        //            isLoadXL_OK = LoadingXL(vOpenExcelFileName, idaXLUPLOAD_ALLOWANCE, 2, 11);

        //            if (isLoadXL_OK == true)
        //            {
        //                vMessage = string.Format("Excel Data Loaded OK [{0}]", vOpenExcelFileName);
        //                isAppInterfaceAdv1.OnAppMessage(vMessage);
        //                System.Windows.Forms.Application.DoEvents();

        //                try
        //                {
        //                    idaXLUPLOAD_ALLOWANCE.Update(); //지급
        //                }
        //                catch
        //                {
        //                    idaXLUPLOAD_ALLOWANCE.Cancel();
        //                }
        //            }
        //            else
        //            {
        //                vMessage = string.Format("Excel Data Loaded Err [{0}]", vOpenExcelFileName);
        //                isAppInterfaceAdv1.OnAppMessage(vMessage);
        //                System.Windows.Forms.Application.DoEvents();
        //            }
        //        }
        //        else if (vIndexTab == 1) //공제
        //        {
        //            idaXLUPLOAD_DEDUCTION.Cancel();

        //            ePath_Deduction.EditValue = vOpenExcelFileName;
        //            isLoadXL_OK = LoadingXL(vOpenExcelFileName, idaXLUPLOAD_DEDUCTION, 2, 9);

        //            if (isLoadXL_OK == true)
        //            {
        //                vMessage = string.Format("Excel Data Loaded OK [{0}]", vOpenExcelFileName);
        //                isAppInterfaceAdv1.OnAppMessage(vMessage);
        //                System.Windows.Forms.Application.DoEvents();

        //                try
        //                {
        //                idaXLUPLOAD_DEDUCTION.Update(); //공제
        //                }
        //                catch
        //                {
        //                    idaXLUPLOAD_ALLOWANCE.Cancel();
        //                }
        //            }
        //            else
        //            {
        //                vMessage = string.Format("Excel Data Loaded Err [{0}]", vOpenExcelFileName);
        //                isAppInterfaceAdv1.OnAppMessage(vMessage);
        //                System.Windows.Forms.Application.DoEvents();
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    this.Cursor = Cursors.Default;
        //    Application.DoEvents();
        //}

        //#endregion;

        //#region ----- Excel Loading Method ----

        //private bool LoadingXL(string pExcelFile, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pRow, int pColumn)
        //{
        //    string vMessage = string.Empty;

        //    bool isLoad_OK = false;

        //    XLoading vImport = null;

        //    try
        //    {
        //        vMessage = string.Format("Excel Data Loading...");
        //        isAppInterfaceAdv1.OnAppMessage(vMessage);
        //        System.Windows.Forms.Application.DoEvents();

        //        vImport = new XLoading(isAppInterfaceAdv1, isMessageAdapter1);

        //        vImport.OpenFileName = pExcelFile;
        //        bool IsOpen = vImport.OpenXL();
        //        if (IsOpen == true)
        //        {
        //            vImport.ReadRow = pRow;        //Excel에서 읽어들일 시작 행
        //            vImport.CountCOLUMN = pColumn; //Excel에서 읽어들일 열 갯수 지정
        //            igbCONDITION.PromptText = string.Format("Inquiry Condition - [Excel Sheet Row : {0}   Column : {1}]", vImport.CountROW, vImport.CountCOLUMN);
        //            System.Windows.Forms.Application.DoEvents();

        //            isLoad_OK = vImport.LoadXL(pAdapter);
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //        System.Windows.Forms.Application.DoEvents();

        //        vImport.DisposeXL();
        //    }

        //    vImport.DisposeXL();

        //    return isLoad_OK;
        //}

        //#endregion;

    }
}