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

namespace HRMF0786
{
    public partial class HRMF0786 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0786()
        {
            InitializeComponent();
        }

        public HRMF0786(Form pMainForm, ISAppInterface pAppInterface)
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
                IDA_OFFICE_TAX_LIST.Fill();
                IGR_OFFICE_TAX_LIST.Focus();
            }
            else if (TB_MAIN.SelectedIndex == 1)
            {
                Search_DB_Detail(OFFICE_TAX_ID.EditValue);
            }
        }

        private void Search_DB_Detail(object pOFFICE_TAX_ID)
        {
            if (iConv.ISNull(pOFFICE_TAX_ID) == string.Empty)
            {
                return;
            }
            IDA_OFFICE_TAX_DOC.OraSelectData.AcceptChanges();
            IDA_OFFICE_TAX_DOC.Refillable = true;

            IDA_OFFICE_TAX_DOC.SetSelectParamValue("P_OFFICE_TAX_ID", pOFFICE_TAX_ID);
            IDA_OFFICE_TAX_DOC.Fill();
            OFFICE_TAX_TYPE_DESC.Focus();
        }

        private Boolean Check_WITHHOLDING_DOC_Added()
        {
            Boolean Row_Added_Status = false;

            for (int r = 0; r < IDA_OFFICE_TAX_DOC.SelectRows.Count; r++)
            {
                if (IDA_OFFICE_TAX_DOC.SelectRows[r].RowState == DataRowState.Added)
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
            HRMF0786_PRINT vHRMF0786_PRINT = new HRMF0786_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vHRMF0786_PRINT.ShowDialog();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (dlgRESULT == DialogResult.OK)
            {
                //인쇄 선택.
                if (vHRMF0786_PRINT.Print_1_YN == "Y")
                {
                    XLPrinting1(pOutput_Type);
                }
                if (vHRMF0786_PRINT.Print_2_YN == "Y")
                {
                    XLPrinting2(pOutput_Type);
                }
            }
            vHRMF0786_PRINT.Dispose();
        }

        private void Set_Pay_Supply_Date()
        {
            IDC_PAY_DATE.SetCommandParamValue("W_WAGE_TYPE", "P1");
            IDC_PAY_DATE.ExecuteNonQuery();
            PAY_SUPPLY_DATE.EditValue = IDC_PAY_DATE.GetCommandParamValue("O_SUPPLY_DATE");
        }

        private void Set_Comp_Tax_Amt()
        {
            IDC_COMP_TAX_AMT.ExecuteNonQuery();
            COMP_TAX_AMT.EditValue = IDC_COMP_TAX_AMT.GetCommandParamValue("O_TAX_AMT");

            TOTAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(COMP_TAX_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(TAX_ADDITION_AMT.EditValue, 0);
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
            IDA_PRINT_OFFICE_TAX_DOC.Fill();
            IDA_PRINT_OFFICE_TAX_DOC_S.Fill();
            IDA_PRINT_SALARY_ITEM.Fill();
            IDA_PRINT_TAXFREE_SALARY.Fill();

            vCountRow = IDA_PRINT_OFFICE_TAX_DOC.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Office_Tax_doc_";

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
                xlPrinting.OpenFileNameExcel = "HRMF0786_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_OFFICE_TAX_DOC, IDA_PRINT_OFFICE_TAX_DOC_S, IDA_PRINT_SALARY_ITEM, IDA_PRINT_TAXFREE_SALARY);

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
            IDA_PRINT_OFFICE_TAX_2.Fill();
            vCountRow = IDA_PRINT_OFFICE_TAX_2.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "OFFICE_TAX_2_";

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
                xlPrinting.OpenFileNameExcel = "HRMF0786_002.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_OFFICE_TAX_2);

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
                    if (IDA_OFFICE_TAX_LIST.IsFocused || IDA_OFFICE_TAX_DOC.IsFocused)
                    {
                        if (Check_WITHHOLDING_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_OFFICE_TAX_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_OFFICE_TAX_DOC.AddOver();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;
                        DUE_DATE.EditValue = SUBMIT_DATE.EditValue;
                        ORIGINAL_DUE_DATE.EditValue = SUBMIT_DATE.EditValue;

                        IDC_GET_OFFICE_TAX_OFFICER_P.ExecuteNonQuery();
                        TAX_OFFICER.EditValue = IDC_GET_OFFICE_TAX_OFFICER_P.GetCommandParamValue("O_TAX_OFFICER");

                        OWNER_TAX_FREE_YN.CheckBoxValue = "N";
                        Set_Pay_Supply_Date();
                        OFFICE_TAX_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_OFFICE_TAX_LIST.IsFocused || IDA_OFFICE_TAX_DOC.IsFocused)
                    {
                        if (Check_WITHHOLDING_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_OFFICE_TAX_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_OFFICE_TAX_DOC.AddUnder();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;
                        DUE_DATE.EditValue = SUBMIT_DATE.EditValue;
                        ORIGINAL_DUE_DATE.EditValue = SUBMIT_DATE.EditValue;

                        IDC_GET_OFFICE_TAX_OFFICER_P.ExecuteNonQuery();
                        TAX_OFFICER.EditValue = IDC_GET_OFFICE_TAX_OFFICER_P.GetCommandParamValue("O_TAX_OFFICER");

                        OWNER_TAX_FREE_YN.CheckBoxValue = "N";
                        Set_Pay_Supply_Date();
                        OFFICE_TAX_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_OFFICE_TAX_LIST.IsFocused)
                    {
                        IDA_OFFICE_TAX_LIST.Update();
                    }
                    else if (IDA_OFFICE_TAX_DOC.IsFocused)
                    {
                        OFFICE_TAX_NO.Focus();
                        IDA_OFFICE_TAX_DOC.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OFFICE_TAX_LIST.IsFocused)
                    {
                        IDA_OFFICE_TAX_LIST.Cancel();
                    }
                    else if (IDA_OFFICE_TAX_DOC.IsFocused)
                    {
                        IDA_OFFICE_TAX_DOC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OFFICE_TAX_LIST.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_OFFICE_TAX_DOC_P.SetCommandParamValue("W_OFFICE_TAX_ID", IGR_OFFICE_TAX_LIST.GetCellValue("OFFICE_TAX_ID"));
                        IDC_DELETE_OFFICE_TAX_DOC_P.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_OFFICE_TAX_DOC_P.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_OFFICE_TAX_DOC_P.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        Search_DB();
                    }
                    else if (IDA_OFFICE_TAX_DOC.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_OFFICE_TAX_DOC_P.SetCommandParamValue("W_OFFICE_TAX_ID", OFFICE_TAX_ID.EditValue);
                        IDC_DELETE_OFFICE_TAX_DOC_P.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_OFFICE_TAX_DOC_P.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_OFFICE_TAX_DOC_P.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        Search_DB_Detail(-1);
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

        private void HRMF0786_Load(object sender, EventArgs e)
        {
            IDA_OFFICE_TAX_DOC.FillSchema();
        }

        private void HRMF0786_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            SUBMIT_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }

        private void IGR_WITHHOLDING_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_OFFICE_TAX_LIST.Row < 1)
            {
                return;
            }
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            Search_DB_Detail(IGR_OFFICE_TAX_LIST.GetCellValue("OFFICE_TAX_ID"));
        }

        private void BTN_PROCESS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //update.
            OFFICE_TAX_NO.Focus();
            IDA_OFFICE_TAX_DOC.Update();

            if (iConv.ISNull(OFFICE_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                OFFICE_TAX_NO.Focus();
                return;
            }

            if (iConv.ISNull(PRE_OFFICE_TAX_ID.EditValue) == String.Empty)
            {
                if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("NFKHRM_10034"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    PRE_OFFICE_TAX_NO.Focus();
                    return;
                } 
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_MAIN_OFFICE_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_MAIN_OFFICE_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_MAIN_OFFICE_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_MAIN_OFFICE_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
         
            // requery.
            Search_DB_Detail(OFFICE_TAX_ID.EditValue);
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(OFFICE_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_CLOSED_OFFICE_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_OFFICE_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_OFFICE_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_OFFICE_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(OFFICE_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;

            IDC_CLOSED_CANCEL_OFFICE_TAX.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_CANCEL_OFFICE_TAX.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_CANCEL_OFFICE_TAX.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_CANCEL_OFFICE_TAX.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BTN_SALARY_DTL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(OFFICE_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                OFFICE_TAX_NO.Focus();
                return;
            }

            DialogResult vResult = DialogResult.None;
            HRMF0786_SALARY vHRMF0786_SALARY = new HRMF0786_SALARY(isAppInterfaceAdv1.AppInterface, OFFICE_TAX_NO.EditValue, OFFICE_TAX_ID.EditValue
                                                                    , STD_YYYYMM.EditValue, PAY_YYYYMM.EditValue, PAY_SUPPLY_DATE.EditValue);
            vResult = vHRMF0786_SALARY.ShowDialog();
            if(vResult == DialogResult.Cancel)
            {
                vHRMF0786_SALARY.Dispose();
                return;
            }
            vHRMF0786_SALARY.Dispose();

            // requery.
            Search_DB_Detail(OFFICE_TAX_ID.EditValue);
        }

        #endregion

        #region ----- 동기화 및 자동 계산 -----

        private void SUBMIT_DATE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            DUE_DATE.EditValue = SUBMIT_DATE.EditValue;
        }

        private void TOTAL_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            PAYMENT_TAX_AMT.EditValue = iConv.ISDecimaltoZero(TOTAL_PAYMENT_AMT.EditValue, 0) -
                                        iConv.ISDecimaltoZero(TAX_FREE_AMT.EditValue, 0);

            Set_Comp_Tax_Amt();
        }

        private void TAX_FREE_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            PAYMENT_TAX_AMT.EditValue = iConv.ISDecimaltoZero(TOTAL_PAYMENT_AMT.EditValue, 0) -
                                        iConv.ISDecimaltoZero(TAX_FREE_AMT.EditValue, 0);

            Set_Comp_Tax_Amt();
        }

        private void COMP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(COMP_TAX_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(TAX_ADDITION_AMT.EditValue, 0);
        }

        private void TAX_ADDITION_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(COMP_TAX_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(TAX_ADDITION_AMT.EditValue, 0);
        }

        private void ORIGINAL_DUE_DATE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //납부지연일수.
            IDC_GET_DELAY_DAY_COUNT.ExecuteNonQuery();
            //가산세액.
            IDC_GET_ADD_TAX_AMT_P.ExecuteNonQuery();
        }

        private void DELAY_DAY_COUNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            //가산세액.
            IDC_GET_ADD_TAX_AMT_P.ExecuteNonQuery();
        }

        private void BAD_PAY_ADDITION_AMT_EditValueChanged(object pSender)
        {
            TAX_ADDITION_AMT.EditValue = iConv.ISDecimaltoZero(BAD_PAY_ADDITION_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(BAD_REPORT_ADDITION_AMT.EditValue, 0);

            TOTAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(COMP_TAX_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(TAX_ADDITION_AMT.EditValue, 0);
        }

        private void BAD_REPORT_ADDITION_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TAX_ADDITION_AMT.EditValue = iConv.ISDecimaltoZero(BAD_PAY_ADDITION_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(BAD_REPORT_ADDITION_AMT.EditValue, 0);

            TOTAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(COMP_TAX_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(TAX_ADDITION_AMT.EditValue, 0);
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

        private void ILA_OFFICE_TAX_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
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
            if (iConv.ISNull(e.Row["OFFICE_TAX_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    }
}