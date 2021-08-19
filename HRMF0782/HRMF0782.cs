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

namespace HRMF0782
{
    public partial class HRMF0782 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0782()
        {
            InitializeComponent();
        }

        public HRMF0782(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        //업체
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
        }

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (TB_MAIN.SelectedIndex == 0)
            {
                IDA_LOCAL_TAX_LIST.Fill();
                IGR_LOCAL_TAX_LIST.Focus();
            }
            else if (TB_MAIN.SelectedIndex == 1)
            {
                Search_DB_Detail(LOCAL_TAX_ID.EditValue);
            }
        }

        private void Search_DB_Detail(object pWITHHOLDING_DOC_ID)
        {
            if (iConv.ISNull(pWITHHOLDING_DOC_ID) == string.Empty)
            {
                return;
            }
            IDA_LOCAL_TAX_DOC.SetSelectParamValue("P_WITHHOLDING_DOC_ID", pWITHHOLDING_DOC_ID);
            IDA_LOCAL_TAX_DOC.Fill();
            LOCAL_TAX_TYPE_DESC.Focus();
        }

        private Boolean Check_LOCAL_TAX_Added()
        {
            Boolean Row_Added_Status = false;

            for (int r = 0; r < IDA_LOCAL_TAX_DOC.SelectRows.Count; r++)
            {
                if (IDA_LOCAL_TAX_DOC.SelectRows[r].RowState == DataRowState.Added)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return (Row_Added_Status);
        }

        private void Printing_DOC(string pOutput_Type)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            HRMF0782_PRINT vHRMF0782_PRINT = new HRMF0782_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vHRMF0782_PRINT.ShowDialog();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (dlgRESULT == DialogResult.OK)
            {
                //인쇄 선택.
                if (vHRMF0782_PRINT.Print_1_YN == "Y")
                {
                    XLPrinting1(pOutput_Type);
                }
                if (vHRMF0782_PRINT.Print_2_YN == "Y")
                {
                    XLPrinting2(pOutput_Type);
                }
            }
            vHRMF0782_PRINT.Dispose();
        }

        private void Set_Pay_Supply_Date()
        {
            IDC_PAY_DATE.SetCommandParamValue("W_WAGE_TYPE", "P1");
            IDC_PAY_DATE.ExecuteNonQuery();
            PAY_SUPPLY_DATE.EditValue = IDC_PAY_DATE.GetCommandParamValue("O_SUPPLY_DATE");
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
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_LOCAL_TAX_DOC.Fill();
            vCountRow = IDA_PRINT_LOCAL_TAX_DOC.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Local_Tax_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0782_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_LOCAL_TAX_DOC);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 2 Methods ----

        private void XLPrinting2(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;
            
            // 데이터 조회.
            IDA_PRINT_LOCAL_TAX_2.Fill();
            vCountRow = IDA_PRINT_LOCAL_TAX_2.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Local_Tax_2_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0782_002.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_LOCAL_TAX_2);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

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
                    if (IDA_LOCAL_TAX_LIST.IsFocused || IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        if (Check_LOCAL_TAX_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_LOCAL_TAX_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_LOCAL_TAX_DOC.AddOver();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;
                        Set_Pay_Supply_Date();
                        LOCAL_TAX_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused || IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        if (Check_LOCAL_TAX_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_LOCAL_TAX_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_LOCAL_TAX_DOC.AddUnder();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;
                        Set_Pay_Supply_Date();
                        LOCAL_TAX_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused)
                    {
                        IDA_LOCAL_TAX_LIST.Update();
                    }
                    else if (IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        LOCAL_TAX_NO.Focus();
                        IDA_LOCAL_TAX_DOC.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused)
                    {
                        IDA_LOCAL_TAX_LIST.Cancel();
                    }
                    else if (IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused)
                    {
                        IDA_LOCAL_TAX_LIST.Delete();
                    }
                    else if (IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    Printing_DOC("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    Printing_DOC("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0782_Load(object sender, EventArgs e)
        {
            IDA_LOCAL_TAX_DOC.FillSchema();
        }

        private void HRMF0782_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            SUBMIT_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }

        private void IGR_LOCAL_TAX_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_LOCAL_TAX_LIST.Row < 1)
            {
                return;
            }
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            Search_DB_Detail(IGR_LOCAL_TAX_LIST.GetCellValue("LOCAL_TAX_ID"));
        }

        private void BTN_PROCESS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //update.
            IDA_LOCAL_TAX_DOC.Update();

            if (iConv.ISNull(LOCAL_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LOCAL_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 귀속년월 및 지급년월이 매년도 2월일 경우 집계데이터 선택 화면 표시.
            object vSET_TYPE = "3";  // 집계데이터 선택 : 1-연말정산데이터, 2-매월징수분 + 연말정산데이터, 3-매월징수분.
            if (iConv.ISNull(STD_YYYYMM.EditValue).Substring(5, 2) == "02" && iConv.ISNull(PAY_YYYYMM.EditValue).Substring(5, 2) == "02")
            {
                DialogResult dlgRESULT;
                HRMF0782_SET vHRMF0782 = new HRMF0782_SET(isAppInterfaceAdv1.AppInterface);
                dlgRESULT =vHRMF0782.ShowDialog();
                if (dlgRESULT == DialogResult.Cancel)
                {
                    return;
                }
                vSET_TYPE = vHRMF0782.Set_Type;
                vHRMF0782.Dispose();
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_MAIN_LOCAL_TAX.SetCommandParamValue("P_SET_TYPE", vSET_TYPE);
            IDC_MAIN_LOCAL_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_MAIN_LOCAL_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_MAIN_LOCAL_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_MAIN_LOCAL_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
         
            // requery.
            Search_DB_Detail(LOCAL_TAX_ID.EditValue);
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(LOCAL_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LOCAL_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_CLOSED_LOCAL_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_LOCAL_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_LOCAL_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_LOCAL_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(LOCAL_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LOCAL_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;

            IDC_CLOSED_CANCEL_LOCAL_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_CANCEL_LOCAL_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_CANCEL_LOCAL_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_CANCEL_LOCAL_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_CALENDAR_YEAR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CALENDAR_YEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today, 1));
        }

        private void ILA_STD_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_STD_YYYYMM_SelectedRowData(object pSender)
        {
            Set_Pay_Supply_Date();
        }
        
        private void ILA_PAY_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_WITHHOLDING_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OFFICE_TAX_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter event ------
        
        private void IDA_WITHHOLDING_DOC_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["LOCAL_TAX_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LOCAL_TAX_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["STD_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(STD_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUBMIT_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SUBMIT_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PAY_SUPPLY_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_SUPPLY_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iConv.ISNull(e.Row["TAX_OFFICER"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(TAX_OFFICER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        
        #endregion

        #region ----- 합계 금액 동기화 -----

        private void A01_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A01_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A01_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A02_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A02_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A02_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A03_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A03_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A03_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A04_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A04_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A04_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A05_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A05_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A05_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A06_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A06_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A06_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A07_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A07_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A07_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A08_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A08_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A08_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A09_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A09_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A09_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A10_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A10_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_STD_TAX_AMT();
        }

        private void A10_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void K10_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void K20_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void K30_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void R10_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void R20_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void R30_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        #endregion

        #region ----- Sum or Total -----
        
        private void TOTAL_PERSON_CNT()
        {
            A90_PERSON_CNT.EditValue = 0;
            A90_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A01_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A03_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A04_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A05_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A07_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A08_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A09_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A10_PERSON_CNT.EditValue);
        }

        private void TOTAL_STD_TAX_AMT()
        {
            A90_STD_TAX_AMT.EditValue = 0;
            A90_STD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A01_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A03_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A04_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A05_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A07_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A08_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A09_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A10_STD_TAX_AMT.EditValue);
        }

        private void TOTAL_LOCAL_TAX_AMT()
        {
            A90_LOCAL_TAX_AMT.EditValue = 0;
            A90_LOCAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A01_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A02_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A03_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A04_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A05_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A06_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A07_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A08_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A09_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A10_LOCAL_TAX_AMT.EditValue);
            CAL_LOCAL_TAX_AMT();
        }
        
        #endregion

        #region ----- 환급액 계산 -----

        //당월조정환급액, 가감세액 및 납부세액 계산.
        private void CAL_LOCAL_TAX_AMT()
        {
            //초기화.
            TOTAL_ADJUST_TAX_AMT.EditValue = 0;
            PAY_LOCAL_TAX_AMT.EditValue = 0;

            //1.중도퇴사자연말정산 환급액 계산
            if (iConv.ISDecimaltoZero(R10_TAX_AMT.EditValue) + iConv.ISDecimaltoZero(R20_TAX_AMT.EditValue) > 0)
            {
                //중도퇴사자 환급액 존재시 지방소득세 합계와 연동해서 처리함
                if (iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) >
                    iConv.ISDecimaltoZero(R10_TAX_AMT.EditValue) +
                    iConv.ISDecimaltoZero(R20_TAX_AMT.EditValue))
                {
                    R30_TAX_AMT.EditValue = iConv.ISDecimaltoZero(R10_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(R20_TAX_AMT.EditValue);
                }
                else if (iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) <=
                        iConv.ISDecimaltoZero(R10_TAX_AMT.EditValue) +
                        iConv.ISDecimaltoZero(R20_TAX_AMT.EditValue))
                {
                    R30_TAX_AMT.EditValue = A90_LOCAL_TAX_AMT.EditValue;
                }
                else
                {
                    R30_TAX_AMT.EditValue = 0;
                }
            }
            else
            {
                R30_TAX_AMT.EditValue = 0;
            }

            // 중도퇴사연말정산환급액 차월이월환급액 계산.
            R40_TAX_AMT.EditValue = iConv.ISDecimaltoZero(R10_TAX_AMT.EditValue) +
                                    iConv.ISDecimaltoZero(R20_TAX_AMT.EditValue) -
                                    iConv.ISDecimaltoZero(R30_TAX_AMT.EditValue);

            //가감세액(조정액) 적용.
            TOTAL_ADJUST_TAX_AMT.EditValue = R30_TAX_AMT.EditValue;

            //2.계속근무자연말정산 환급액 계산
            if (iConv.ISDecimaltoZero(K10_TAX_AMT.EditValue) + iConv.ISDecimaltoZero(K20_TAX_AMT.EditValue) > 0)
            {
                //계속근무자연말정산 환급액 존재시 지방소득세 합계와 연동해서 처리함
                if ((iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
                    iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue)) >
                    iConv.ISDecimaltoZero(K10_TAX_AMT.EditValue) +
                    iConv.ISDecimaltoZero(K20_TAX_AMT.EditValue))
                {
                    K30_TAX_AMT.EditValue = iConv.ISDecimaltoZero(K10_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(K20_TAX_AMT.EditValue);
                }
                else if ((iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
                        iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue)) <=
                        iConv.ISDecimaltoZero(K10_TAX_AMT.EditValue) +
                        iConv.ISDecimaltoZero(K20_TAX_AMT.EditValue))
                {
                    K30_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue);
                }
            }
            else
            {
                K30_TAX_AMT.EditValue = 0;
            }


            // 계속근무자연말정산환급액 차월이월환급액 계산.
            K40_TAX_AMT.EditValue = iConv.ISDecimaltoZero(K10_TAX_AMT.EditValue) +
                                    iConv.ISDecimaltoZero(K20_TAX_AMT.EditValue) -
                                    iConv.ISDecimaltoZero(K30_TAX_AMT.EditValue);

            //가감세액(조정액) 계산.
            TOTAL_ADJUST_TAX_AMT.EditValue = iConv.ISDecimaltoZero(K30_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(R30_TAX_AMT.EditValue);


            //납부세액 계산.
            PAY_LOCAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue);
        }

        #endregion

    }
}