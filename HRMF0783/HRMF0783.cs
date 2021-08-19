using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.Runtime.InteropServices;       //호환되지 않은DLL을 사용할 때.

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0783
{
    public partial class HRMF0783 : Office2007Form
    {
        #region ----- API Dll Import -----

        [DllImport("fcrypt_es.dll")]
        extern public static int DSFC_EncryptFile(int hWnd, string pszPlainFilePathName, string pszEncFilePathName, string pszPassword, uint nOption);

        string inputPath;
        string OutputPath;
        string Password;
        uint DSFC_OPT_OVERWRITE_OUTPUT;
        int nRet;

        #endregion;

        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0783()
        {
            InitializeComponent();
        }

        public HRMF0783(Form pMainForm, ISAppInterface pAppInterface)
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

            IDA_LOCAL_TAX_DOC.OraSelectData.AcceptChanges();
            IDA_LOCAL_TAX_DOC.Refillable = true;

            IDA_LOCAL_TAX_DOC.SetSelectParamValue("P_LOCAL_TAX_ID", pWITHHOLDING_DOC_ID);
            IDA_LOCAL_TAX_DOC.Fill();
            if(iConv.ISNull(LOCAL_TAX_PAYMENT_TYPE.EditValue) == "2")
            {
                RB_HALF.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            else
            {
                RB_MONTH.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
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
            HRMF0783_PRINT vHRMF0783_PRINT = new HRMF0783_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vHRMF0783_PRINT.ShowDialog();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (dlgRESULT == DialogResult.OK)
            {
                //인쇄 선택.
                if (vHRMF0783_PRINT.Print_1_YN == "Y")
                {
                    XLPrinting1(pOutput_Type);
                }
                if (vHRMF0783_PRINT.Print_2_YN == "Y")
                {
                    XLPrinting2(pOutput_Type);
                }
            }
            vHRMF0783_PRINT.Dispose();
        }

        private void Set_Pay_Supply_Date()
        {
            IDC_PAY_DATE.SetCommandParamValue("W_WAGE_TYPE", "P1");
            IDC_PAY_DATE.ExecuteNonQuery();
            PAY_SUPPLY_DATE.EditValue = IDC_PAY_DATE.GetCommandParamValue("O_SUPPLY_DATE");
        }


        private string EXPORT_VALIDATE()
        {
            string vRETURN = "N";
             
            if (iConv.ISNull(LOCAL_TAX_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(LOCAL_TAX_NO)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(STD_YYYYMM)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(PAY_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(PAY_YYYYMM)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(SUBMIT_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(SUBMIT_DATE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(ORI_PAYMENT_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(ORI_PAYMENT_DATE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(PAYMENT_DUE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(PAYMENT_DUE_DATE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            } 
            vRETURN = "Y";
            return vRETURN;
        } 

        private void Button_Control(string pEnabled_YN)
        {
            if (pEnabled_YN == "Y")
            {
                BTN_PROCESS.Enabled = true;
                BTN_CLOSED_OK.Enabled = true;
                BTN_CLOSED_CANCEL.Enabled = true;
                BTN_FILE.Enabled = true;
            }
            else
            {
                BTN_PROCESS.Enabled = false;
                BTN_CLOSED_OK.Enabled = false;
                BTN_CLOSED_CANCEL.Enabled = false;
                BTN_FILE.Enabled = false;
            }
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
                xlPrinting.OpenFileNameExcel = "HRMF0783_001.xlsx";
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
                xlPrinting.OpenFileNameExcel = "HRMF0783_002.xlsx";
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
                        RB_MONTH.CheckedState = ISUtil.Enum.CheckedState.Checked;

                        IDC_GET_LOCAL_OFFICER_P.ExecuteNonQuery();
                        TAX_OFFICER.EditValue = IDC_GET_LOCAL_OFFICER_P.GetCommandParamValue("O_LOCAL_OFFICER");

                        Set_Pay_Supply_Date();
                        Init_PAYMENT_DATE(PAY_YYYYMM.EditValue);
                        LOCAL_TAX_TYPE_DESC.Focus();
                    }
                    else if(IDA_LOCAL_TAX_DOC_ITEM.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC_ITEM.AddOver();
                        IGR_LOCAL_TAX_DOC_ITEM.Focus();
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
                        RB_MONTH.CheckedState = ISUtil.Enum.CheckedState.Checked;

                        IDC_GET_LOCAL_OFFICER_P.ExecuteNonQuery();
                        TAX_OFFICER.EditValue = IDC_GET_LOCAL_OFFICER_P.GetCommandParamValue("O_LOCAL_OFFICER");

                        Set_Pay_Supply_Date();
                        Init_PAYMENT_DATE(PAY_YYYYMM.EditValue);
                        LOCAL_TAX_TYPE_DESC.Focus();
                    }
                    else if (IDA_LOCAL_TAX_DOC_ITEM.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC_ITEM.AddUnder();
                        IGR_LOCAL_TAX_DOC_ITEM.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused)
                    {
                        IDA_LOCAL_TAX_LIST.Update();
                    }
                    else 
                    {
                        LOCAL_TAX_NO.Focus();
                        CAL_LOCAL_TAX_AMT();
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
                        IDA_LOCAL_TAX_DOC_ITEM.Cancel();
                        IDA_LOCAL_TAX_DOC.Cancel();
                    }
                    else if (IDA_LOCAL_TAX_DOC_ITEM.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC_ITEM.Cancel(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_LOCAL_TAX_LIST.IsFocused)
                    {
                        if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_LOCAL_TAX_DOC.SetCommandParamValue("W_LOCAL_TAX_ID", IGR_LOCAL_TAX_LIST.GetCellValue("LOCAL_TAX_ID"));
                        IDC_DELETE_LOCAL_TAX_DOC.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_LOCAL_TAX_DOC.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_LOCAL_TAX_DOC.GetCommandParamValue("O_MESSAGE"));
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
                    else if (IDA_LOCAL_TAX_DOC.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_LOCAL_TAX_DOC.SetCommandParamValue("W_LOCAL_TAX_ID", LOCAL_TAX_ID.EditValue);
                        IDC_DELETE_LOCAL_TAX_DOC.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_LOCAL_TAX_DOC.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_LOCAL_TAX_DOC.GetCommandParamValue("O_MESSAGE"));
                        if(vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        Search_DB_Detail(-1);
                    }
                    else if (IDA_LOCAL_TAX_DOC_ITEM.IsFocused)
                    {
                        IDA_LOCAL_TAX_DOC_ITEM.Delete();
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

        private void HRMF0783_Load(object sender, EventArgs e)
        {
            IDA_LOCAL_TAX_DOC.FillSchema();
        }

        private void HRMF0783_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            SUBMIT_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }

        private void RB_MONTH_Click(object sender, EventArgs e)
        {
            if(RB_MONTH.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                LOCAL_TAX_PAYMENT_TYPE.EditValue = RB_MONTH.RadioCheckedString;
            }
        }

        private void RB_HALF_Click(object sender, EventArgs e)
        {
            if (RB_HALF.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                LOCAL_TAX_PAYMENT_TYPE.EditValue = RB_HALF.RadioCheckedString;
            }
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
            if (iConv.ISNull(PAY_YYYYMM.EditValue).Substring(5, 2) == "02")
            {
                DialogResult dlgRESULT;
                HRMF0783_SET vHRMF0783 = new HRMF0783_SET(isAppInterfaceAdv1.AppInterface);
                dlgRESULT =vHRMF0783.ShowDialog();
                if (dlgRESULT == DialogResult.Cancel)
                {
                    return;
                }
                vSET_TYPE = vHRMF0783.Set_Type;
                vHRMF0783.Dispose();
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

        private void BTN_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(LOCAL_TAX_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(LOCAL_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            Button_Control("N");  //버튼 사용 불가 만들기. 

            //전산매체 작성.
            Export_File(); 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Init_PAYMENT_DATE(object pPAY_YYYYMM)
        {
            IDC_GET_PAYMENT_DATE_P.SetCommandParamValue("P_PAY_YYYYMM", pPAY_YYYYMM);
            IDC_GET_PAYMENT_DATE_P.ExecuteNonQuery();

            SUBMIT_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_SUBMIT_DATE");
            ORI_PAYMENT_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_ORI_PAYMENT_DATE");
            PAYMENT_DUE_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_PAYMENT_DUE_DATE");
        }

        private decimal Get_Local_Tax_Amt(string pLocal_Tax_Type, object pComp_Std_Amt)
        {
            IDC_GET_LOCAL_TAX_P.SetCommandParamValue("P_LOCAL_TAX_TYPE", pLocal_Tax_Type);
            IDC_GET_LOCAL_TAX_P.SetCommandParamValue("P_COMP_STD_AMT", pComp_Std_Amt);
            IDC_GET_LOCAL_TAX_P.ExecuteNonQuery();
            decimal vLocal_Tax_Amt = iConv.ISDecimaltoZero(IDC_GET_LOCAL_TAX_P.GetCommandParamValue("O_LOCAL_TAX_AMT"));
            return vLocal_Tax_Amt; 
        }

        private void Init_ADD_PAYMENT_DATE(object pPAY_YYYYMM)
        {
            IDC_GET_PAYMENT_DATE_P.SetCommandParamValue("P_PAY_YYYYMM", pPAY_YYYYMM);
            IDC_GET_PAYMENT_DATE_P.ExecuteNonQuery();

            ADD_ORI_PAYMENT_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_ORI_PAYMENT_DATE");
            ADD_PAYMENT_DUE_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_PAYMENT_DUE_DATE");
        }

        private void IGR_LOCAL_TAX_DOC_ITEM_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_COMP_TAX_AMT = IGR_LOCAL_TAX_DOC_ITEM.GetColumnToIndex("COMP_TAX_AMT");
            int vIDX_LOCAL_TAX_AMT = IGR_LOCAL_TAX_DOC_ITEM.GetColumnToIndex("LOCAL_TAX_AMT");
            int vIDX_ADJUST_TAX_AMT = IGR_LOCAL_TAX_DOC_ITEM.GetColumnToIndex("ADJUST_TAX_AMT");
            string vLOCAL_TAX_ITEM = iConv.ISNull(IGR_LOCAL_TAX_DOC_ITEM.GetCellValue("LOCAL_TAX_ITEM"));
            if (vIDX_COMP_TAX_AMT == e.ColIndex)
            {
                decimal vLOCAL_TAX_AMT = Get_Local_Tax_Amt(vLOCAL_TAX_ITEM, e.NewValue);
                IGR_LOCAL_TAX_DOC_ITEM.SetCellValue("LOCAL_TAX_AMT", vLOCAL_TAX_AMT);

                Init_Local_Item_Amt(vLOCAL_TAX_AMT, IGR_LOCAL_TAX_DOC_ITEM.GetCellValue("ADJUST_TAX_AMT"));
            }
            else if (vIDX_LOCAL_TAX_AMT == e.ColIndex)
            {
                Init_Local_Item_Amt(e.NewValue, IGR_LOCAL_TAX_DOC_ITEM.GetCellValue("ADJUST_TAX_AMT"));
            }
            else if (vIDX_ADJUST_TAX_AMT == e.ColIndex)
            { 
                Init_Local_Item_Amt(IGR_LOCAL_TAX_DOC_ITEM.GetCellValue("ADJUST_TAX_AMT"), e.NewValue);
            }
        }

        private void Init_Local_Item_Amt(object pLocal_Tax_Amt, object pAdjust_Tax_Amt)
        {
            decimal vFix_Local_Tax_Amt = iConv.ISDecimaltoZero(pLocal_Tax_Amt, 0) + iConv.ISDecimaltoZero(pAdjust_Tax_Amt, 0);
            vFix_Local_Tax_Amt = Math.Truncate(vFix_Local_Tax_Amt / 10) * 10;
            IGR_LOCAL_TAX_DOC_ITEM.SetCellValue("FIX_LOCAL_TAX_AMT", vFix_Local_Tax_Amt); 
        }

        #endregion


        #region ----- Export TXT File ------

        private void Export_File()
        {
            //전산매체 암호화 암호 입력 받기.
            DialogResult vdlgResult;
            object vENCRYPT_PASSWORD = String.Empty;
            HRMF0783_FILE vHRMF0783_FILE = new HRMF0783_FILE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0783_FILE.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
                vENCRYPT_PASSWORD = vHRMF0783_FILE.Get_Encrypt_Password;
            }

            if (iConv.ISNull(vENCRYPT_PASSWORD) == string.Empty)
            {
                Button_Control("Y");  //버튼 사용 불가 만들기. 
                return;
            }

            Button_Control("N");  //버튼 사용 불가 만들기.
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            IDA_LOCAL_TAX_FILE.Fill();
            if (IDA_LOCAL_TAX_FILE.SelectRows.Count < 1)
            {
                Button_Control("Y");  //버튼 사용 불가 만들기.
                MessageBoxAdv.Show("Not found Data. Fail, Export file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isAppInterfaceAdv1.OnAppMessage("Export File start...");

            string vSaveFile_name = string.Empty;
            string vFileName = string.Empty;
            string vFileExted = string.Empty;
            string vFilePath = "C:\\ersdata";

            int euckrCodepage = 51949;
            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

            //파일명(제출일자).
            vFileName = iDate.ISGetDate(SUBMIT_DATE.EditValue).ToShortDateString().Replace("-", "");

            //신고구분에 따른 확장자 지정.
            vFileExted = "A103900.1";
            vFileName = string.Format("{0}{1}", vFileName, vFileExted);

            //파일 경로 디렉토리 존재 여부 체크(없으면 생성).
            if (System.IO.Directory.Exists(vFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(vFilePath);
            }

            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = vFileName;
            saveFileDialog1.DefaultExt = ".1";  // String.Format(".{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            //System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(vFilePath);
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Text Files (*.1)|*.1";//String.Format("Text Files (*.{0})|*.{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                Application.DoEvents();

                vSaveFile_name = saveFileDialog1.FileName;
                //기존 동일한 파일 삭제.
                if (System.IO.File.Exists(vSaveFile_name) == true)
                {
                    try
                    {
                        System.IO.File.Delete(vSaveFile_name);
                    }
                    catch (Exception ex)
                    {
                        MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                //파일 생성.
                try
                {
                    vWriteFile = System.IO.File.Open(vSaveFile_name, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    foreach (DataRow pRow in IDA_LOCAL_TAX_FILE.OraSelectData.Rows)
                    {
                        vSaveString = new System.Text.StringBuilder();
                        vSaveString.Append(pRow["REPORT_FILE"]);
                        vSaveString.Append("\r\n");

                        System.Text.Encoding vEuckr = System.Text.Encoding.GetEncoding(euckrCodepage);
                        byte[] vSavebytes = vEuckr.GetBytes(vSaveString.ToString());

                        int vSaveStringLength = vSavebytes.Length;
                        vWriteFile.Write(vSavebytes, 0, vSaveStringLength);
                    } 
                }
                catch (Exception ex)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                vWriteFile.Dispose();

                nRet = 0;
                inputPath = vSaveFile_name;// "20120410.201";//pFileName;
                OutputPath = string.Format("{0}.erc", vSaveFile_name);
                Password = vENCRYPT_PASSWORD.ToString();
                DSFC_OPT_OVERWRITE_OUTPUT = 1;
                nRet = DSFC_EncryptFile(0, inputPath, OutputPath, Password, DSFC_OPT_OVERWRITE_OUTPUT);
                if (nRet != 0)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    Application.DoEvents();
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    MessageBox.Show(String.Format("Encrypt Error : {0}", nRet), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                System.IO.File.Delete(vSaveFile_name);
                System.IO.File.Copy(OutputPath, inputPath, true);
                System.IO.File.Delete(OutputPath);
            }
            Button_Control("Y");  //버튼 사용 만들기.
            isAppInterfaceAdv1.OnAppMessage("Complete, Export file");
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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

        private void ILA_PAY_YYYYMM_SelectedRowData(object pSender)
        {
            Init_PAYMENT_DATE(PAY_YYYYMM.EditValue);
        }

        private void ILA_ADD_PAYMENTS_MONTH_SelectedRowData(object pSender)
        {
            Init_ADD_PAYMENT_DATE(ADD_PAYMENTS_MONTH.EditValue);
            CAL_ADD_ADDITION_LOCAL_TAX_AMT("ADD");
        }

        private void ILA_WITHHOLDING_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OFFICE_TAX_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_LOCAL_TAX_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "LOCAL_TAX_ITEM");
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
            A01_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A01", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A02_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A02", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A03_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A03", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A04_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A04", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A05_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A05", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A06_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A06", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A07_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A07", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A07_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A11_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A11_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            A11_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A11", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A11_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A09_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A09_STD_TAX_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            A09_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A09", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
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
            A10_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A10", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A10_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A12_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PERSON_CNT();
        }

        private void A12_STD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            A12_LOCAL_TAX_AMT.EditValue = Get_Local_Tax_Amt("A12", e.EditValue);
            TOTAL_STD_TAX_AMT();
            TOTAL_LOCAL_TAX_AMT();
        }

        private void A12_LOCAL_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_LOCAL_TAX_AMT();
        }

        private void THIS_ETC_REFUND_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void YEAR_ADJ_REFUND_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void RETIRE_ADJ_REFUND_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void THIS_ETC_ADD_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void YEAR_ADJ_ADD_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_LOCAL_TAX_AMT();
        }

        private void ADD_PAYMENTS_AMT_CurrentEditValidating_1(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_ADD_ADDITION_LOCAL_TAX_AMT("ADD");
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
                                        iConv.ISDecimaltoZero(A11_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A09_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A10_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A12_PERSON_CNT.EditValue);
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
                                        iConv.ISDecimaltoZero(A11_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A09_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A10_STD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A12_STD_TAX_AMT.EditValue);
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
                                            iConv.ISDecimaltoZero(A11_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A09_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A10_LOCAL_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A12_LOCAL_TAX_AMT.EditValue);

            CAL_ADDITION_LOCAL_TAX_AMT();
            CAL_LOCAL_TAX_AMT();
        }

        #endregion

        #region ----- 가산세 계산 -----

        //당월조정환급액, 가감세액 및 납부세액 계산.
        private void CAL_ADDITION_LOCAL_TAX_AMT()
        {
            //초기화.
            FIX_ADD_LOCAL_TAX_AMT.EditValue = 0;
            if (iDate.ISGetDate(ORI_PAYMENT_DATE.EditValue) < iDate.ISGetDate(PAYMENT_DUE_DATE.EditValue))
            {
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_PAY_YYYYMM", PAY_YYYYMM.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_ADD_LOCAL_TAX_TYPE", "FIX");
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_ORI_PAYMENT_DATE", ORI_PAYMENT_DATE.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_PAYMENT_DUE_DATE", PAYMENT_DUE_DATE.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_COMP_STD_AMT", A90_LOCAL_TAX_AMT.EditValue);
                IDC_ADD_LOCAL_TAX_P.ExecuteNonQuery();
                FIX_ADD_LOCAL_TAX_AMT.EditValue = IDC_ADD_LOCAL_TAX_P.GetCommandParamValue("O_LOCAL_TAX_AMT");
            } 
        }

        //당월조정환급액, 가감세액 및 납부세액 계산.
        private void CAL_ADD_ADDITION_LOCAL_TAX_AMT(object pADD_LOCAL_TAX_TYPE)
        {
            //초기화.
            ADD_LOCAL_TAX_AMT.EditValue = 0;
            Init_ADD_PAYMENT_DATE(ADD_PAYMENTS_MONTH.EditValue);
            if (iDate.ISGetDate(ADD_ORI_PAYMENT_DATE.EditValue) < iDate.ISGetDate(ADD_PAYMENT_DUE_DATE.EditValue))
            {
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_PAY_YYYYMM", ADD_PAYMENTS_MONTH.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_ADD_LOCAL_TAX_TYPE", pADD_LOCAL_TAX_TYPE);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_ORI_PAYMENT_DATE", ADD_ORI_PAYMENT_DATE.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_PAYMENT_DUE_DATE", ADD_PAYMENT_DUE_DATE.EditValue);
                IDC_ADD_LOCAL_TAX_P.SetCommandParamValue("P_COMP_STD_AMT", ADD_PAYMENTS_AMT.EditValue);
                IDC_ADD_LOCAL_TAX_P.ExecuteNonQuery();
                ADD_LOCAL_TAX_AMT.EditValue = IDC_ADD_LOCAL_TAX_P.GetCommandParamValue("O_LOCAL_TAX_AMT");
            }
        }

        #endregion
          
        #region ----- 합계 금액 계산 -----

        //당월조정환급액, 가감세액 및 납부세액 계산.
        private void CAL_LOCAL_TAX_AMT()
        {
            //초기화.
            decimal vTEMP_AMT = 0;
            REFUND_SUM_AMT.EditValue = 0;
            ADD_SUM_AMT.EditValue = 0;
            TOTAL_ADJUST_TAX_AMT.EditValue = 0;
            REMAIN_LOCAL_TAX_AMT.EditValue = 0;
            PAY_LOCAL_TAX_AMT.EditValue = 0;

            //환급 합계.
            REFUND_SUM_AMT.EditValue = iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(RETIRE_ADJ_REFUND_AMT.EditValue, 0);

            //추가 납부합계.
            ADD_SUM_AMT.EditValue = iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(ADD_PAYMENTS_AMT.EditValue, 0) +
                                        iConv.ISDecimaltoZero(ADD_LOCAL_TAX_AMT.EditValue, 0);


            //가감합계.
            TOTAL_ADJUST_TAX_AMT.EditValue = iConv.ISDecimaltoZero(ADD_SUM_AMT.EditValue, 0) -
                                                iConv.ISDecimaltoZero(REFUND_SUM_AMT.EditValue, 0);

            //차감후 환급금액.
            vTEMP_AMT = 0;
            vTEMP_AMT = iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue, 0) +
                        iConv.ISDecimaltoZero(FIX_ADD_LOCAL_TAX_AMT.EditValue, 0) +
                        iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue, 0);
            if (vTEMP_AMT < 0)
            {
                REMAIN_LOCAL_TAX_AMT.EditValue = Math.Abs(vTEMP_AMT);
                PAY_LOCAL_TAX_AMT.EditValue = 0;
            }
            else
            {
                PAY_LOCAL_TAX_AMT.EditValue = vTEMP_AMT;
            }

                                                


            ////1.중도퇴사자연말정산 환급액 계산
            //if (iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue) + iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue) > 0)
            //{
            //    //중도퇴사자 환급액 존재시 지방소득세 합계와 연동해서 처리함
            //    if (iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) >
            //        iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue) +
            //        iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue))
            //    {
            //        ADD_PAYMENTS_MONTH.EditValue = iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue) +
            //                                iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue);
            //    }
            //    else if (iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) <=
            //            iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue) +
            //            iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue))
            //    {
            //        ADD_PAYMENTS_MONTH.EditValue = A90_LOCAL_TAX_AMT.EditValue;
            //    }
            //    else
            //    {
            //        ADD_PAYMENTS_MONTH.EditValue = 0;
            //    }
            //}
            //else
            //{
            //    ADD_PAYMENTS_MONTH.EditValue = 0;
            //}

            //// 중도퇴사연말정산환급액 차월이월환급액 계산.
            //R40_TAX_AMT.EditValue = iConv.ISDecimaltoZero(THIS_ETC_ADD_AMT.EditValue) +
            //                        iConv.ISDecimaltoZero(YEAR_ADJ_ADD_AMT.EditValue) -
            //                        iConv.ISDecimaltoZero(ADD_PAYMENTS_MONTH.EditValue);

            ////가감세액(조정액) 적용.
            //TOTAL_ADJUST_TAX_AMT.EditValue = ADD_PAYMENTS_MONTH.EditValue;

            ////2.계속근무자연말정산 환급액 계산
            //if (iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue) + iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue) > 0)
            //{
            //    //계속근무자연말정산 환급액 존재시 지방소득세 합계와 연동해서 처리함
            //    if ((iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
            //        iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue)) >
            //        iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue) +
            //        iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue))
            //    {
            //        RETIRE_ADJ_REFUND_AMT.EditValue = iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue) +
            //                                iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue);
            //    }
            //    else if ((iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
            //            iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue)) <=
            //            iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue) +
            //            iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue))
            //    {
            //        RETIRE_ADJ_REFUND_AMT.EditValue = iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
            //                                iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue);
            //    }
            //}
            //else
            //{
            //    RETIRE_ADJ_REFUND_AMT.EditValue = 0;
            //}


            //// 계속근무자연말정산환급액 차월이월환급액 계산.
            //REFUND_SUM_AMT.EditValue = iConv.ISDecimaltoZero(THIS_ETC_REFUND_AMT.EditValue) +
            //                        iConv.ISDecimaltoZero(YEAR_ADJ_REFUND_AMT.EditValue) -
            //                        iConv.ISDecimaltoZero(RETIRE_ADJ_REFUND_AMT.EditValue);

            ////가감세액(조정액) 계산.
            //TOTAL_ADJUST_TAX_AMT.EditValue = iConv.ISDecimaltoZero(RETIRE_ADJ_REFUND_AMT.EditValue) +
            //                                    iConv.ISDecimaltoZero(ADD_PAYMENTS_MONTH.EditValue);


            ////납부세액 계산.
            //PAY_LOCAL_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_LOCAL_TAX_AMT.EditValue) -
            //                                iConv.ISDecimaltoZero(TOTAL_ADJUST_TAX_AMT.EditValue);
        }

        #endregion

    }
}