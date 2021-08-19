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

namespace HRMF0789
{
    public partial class HRMF0789 : Office2007Form
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

        public HRMF0789()
        {
            InitializeComponent();
        }

        public HRMF0789(Form pMainForm, ISAppInterface pAppInterface)
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
                IDA_SLC_DOC_LIST.Fill();
                IGR_SLC_DOC_LIST.Focus();
            }
            else if (TB_MAIN.SelectedIndex == 1)
            {
                Search_DB_Detail(SLC_DOC_ID.EditValue);
            }
        }

        private void Search_DB_Detail(object pSLC_DOC_ID)
        {
            if (iConv.ISNull(pSLC_DOC_ID) == string.Empty)
            {
                return;
            }

            IDA_SLC_DOC.OraSelectData.AcceptChanges();
            IDA_SLC_DOC.Refillable = true;

            IGR_SLC_DOC_ITEM_B.LastConfirmChanges();
            IDA_SLC_DOC_ITEM_B.OraSelectData.AcceptChanges();
            IDA_SLC_DOC_ITEM_B.Refillable = true;

            IDA_SLC_DOC.SetSelectParamValue("P_SLC_DOC_ID", pSLC_DOC_ID);
            IDA_SLC_DOC.Fill();
            if(iConv.ISNull(PAYMENT_ALL_YN.EditValue) == "N")
            {
                RB_PAYMENT_ALL_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            else
            {
                RB_PAYMENT_ALL_Y.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            SLC_DOC_TYPE_DESC.Focus();
        }

        private Boolean Check_SLC_DOC_Added()
        {
            Boolean Row_Added_Status = false;

            for (int r = 0; r < IDA_SLC_DOC.SelectRows.Count; r++)
            {
                if (IDA_SLC_DOC.SelectRows[r].RowState == DataRowState.Added)
                {
                    Row_Added_Status = true;
                }
            }
            for (int r = 0; r < IDA_SLC_DOC_ITEM_B.SelectRows.Count; r++)
            {
                if (IDA_SLC_DOC_ITEM_B.SelectRows[r].RowState == DataRowState.Added)
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

        private void Insert_DB()
        {
            CORP_ID.EditValue = CORP_ID_0.EditValue;
            CORP_NAME.EditValue = CORP_NAME_0.EditValue;
            STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            Init_PAYMENT_DATE(STD_YYYYMM.EditValue);
            PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            SUBMIT_DATE.EditValue = DateTime.Today;
            RB_PAYMENT_ALL_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            PAYMENT_ALL_YN.EditValue = RB_PAYMENT_ALL_N.RadioCheckedString;
            Set_Pay_Supply_Date(); 

            SLC_DOC_TYPE_DESC.Focus();
        }

        private void Printing_DOC(string pOutput_Type)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
             
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            XLPrinting1(pOutput_Type);
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
             
            if (iConv.ISNull(SLC_DOC_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(SLC_DOC_NO)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            IDA_PRINT_SLC_DOC.Fill();
            IDA_PRINT_SLC_DOC_ITEM_A.Fill();
            IDA_PRINT_SLC_DOC_ITEM_B.Fill();
            vCountRow = IDA_PRINT_SLC_DOC.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "SLC_";

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

            //인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0789_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_SLC_DOC, IDA_PRINT_SLC_DOC_ITEM_A, IDA_PRINT_SLC_DOC_ITEM_B);

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
                    if (IDA_SLC_DOC_LIST.IsFocused || IDA_SLC_DOC.IsFocused)
                    {
                        if (Check_SLC_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_SLC_DOC_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_SLC_DOC.AddOver();
                        Insert_DB();
                    }
                    else if(IDA_SLC_DOC_ITEM_B.IsFocused)
                    {
                        IDA_SLC_DOC_ITEM_B.AddOver();
                        IGR_SLC_DOC_ITEM_B.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_SLC_DOC_LIST.IsFocused || IDA_SLC_DOC.IsFocused)
                    {
                        if (Check_SLC_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_SLC_DOC_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_SLC_DOC.AddUnder();
                        Insert_DB(); 
                    }
                    else if (IDA_SLC_DOC_ITEM_B.IsFocused)
                    {
                        IDA_SLC_DOC_ITEM_B.AddUnder();
                        IGR_SLC_DOC_ITEM_B.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_SLC_DOC_LIST.IsFocused)
                    {
                        IDA_SLC_DOC_LIST.Update();
                    }
                    else 
                    {
                        SLC_DOC_NO.Focus(); 
                        IDA_SLC_DOC.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SLC_DOC_LIST.IsFocused)
                    {
                        IDA_SLC_DOC_LIST.Cancel();
                    }
                    else if (IDA_SLC_DOC.IsFocused)
                    {
                        IDA_SLC_DOC_ITEM_B.Cancel();
                        IDA_SLC_DOC.Cancel();
                    }
                    else if (IDA_SLC_DOC_ITEM_B.IsFocused)
                    {
                        IDA_SLC_DOC_ITEM_B.Cancel(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_SLC_DOC_LIST.IsFocused)
                    {
                        if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDA_SLC_DOC_LIST.Delete();
                    }
                    else if (IDA_SLC_DOC.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_DELETE_SLC_DOC_P.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_SLC_DOC_P.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_SLC_DOC_P.GetCommandParamValue("O_MESSAGE"));
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
                    else if (IDA_SLC_DOC_ITEM_B.IsFocused)
                    {
                        IDA_SLC_DOC_ITEM_B.Delete();
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

        private void HRMF0789_Load(object sender, EventArgs e)
        {
            IDA_SLC_DOC.FillSchema();
        }

        private void HRMF0789_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            SUBMIT_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }

        private void RB_PAYMENT_ALL_Y_Click(object sender, EventArgs e)
        {
            if(RB_PAYMENT_ALL_Y.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                PAYMENT_ALL_YN.EditValue = RB_PAYMENT_ALL_Y.RadioCheckedString;
            }
        }

        private void RB_PAYMENT_ALL_N_Click(object sender, EventArgs e)
        {
            if (RB_PAYMENT_ALL_N.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                PAYMENT_ALL_YN.EditValue = RB_PAYMENT_ALL_N.RadioCheckedString;
            }
        }

        private void IGR_LOCAL_TAX_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_SLC_DOC_LIST.Row < 1)
            {
                return;
            }
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            Search_DB_Detail(IGR_SLC_DOC_LIST.GetCellValue("SLC_DOC_ID"));
        }

        private void BTN_PROCESS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //update.
            SLC_DOC_NO.Focus();
            IDA_SLC_DOC.Update();

            if (iConv.ISNull(SLC_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLC_DOC_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
             
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null; 
            IDC_MAIN_SLC_DOC.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_MAIN_SLC_DOC.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_MAIN_SLC_DOC.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_MAIN_SLC_DOC.ExcuteError || vSTATUS == "F")
            {
                if (vSTATUS != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
         
            // requery.
            Search_DB_Detail(SLC_DOC_ID.EditValue);
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SLC_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLC_DOC_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_CLOSED_SLC_DOC.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_SLC_DOC.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_SLC_DOC.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_SLC_DOC.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            Search_DB_Detail(SLC_DOC_ID.EditValue);
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SLC_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLC_DOC_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;

            IDC_CLOSED_CANCEL_SLC_DOC.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_CANCEL_SLC_DOC.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_CANCEL_SLC_DOC.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_CANCEL_SLC_DOC.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;                
            }
            Search_DB_Detail(SLC_DOC_ID.EditValue);
        }

        private void BTN_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(SLC_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLC_DOC_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

            PAY_YYYYMM.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_PAY_YYYYMM");
            SUBMIT_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_SUBMIT_DATE");
            PAY_SUPPLY_DATE.EditValue = IDC_GET_PAYMENT_DATE_P.GetCommandParamValue("O_PAY_SUPPLY_DATE"); 
        } 

        #endregion


        #region ----- Export TXT File ------

        private void Export_File()
        {
            //전산매체 암호화 암호 입력 받기.
            DialogResult vdlgResult;
            object vENCRYPT_PASSWORD = String.Empty;
            HRMF0789_FILE vHRMF0789_FILE = new HRMF0789_FILE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0789_FILE.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
                vENCRYPT_PASSWORD = vHRMF0789_FILE.Get_Encrypt_Password;
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

            IDA_SLC_DOC_FILE.Fill();
            if (IDA_SLC_DOC_FILE.SelectRows.Count < 1)
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
            vFileExted = "800";
            vFileName = string.Format("{0}.{1}", vFileName, vFileExted);

            //파일 경로 디렉토리 존재 여부 체크(없으면 생성).
            if (System.IO.Directory.Exists(vFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(vFilePath);
            }

            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = vFileName;
            saveFileDialog1.DefaultExt = ".*";  // String.Format(".{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            //System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(vFilePath);
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Text Files (*.*)|*.*";//String.Format("Text Files (*.{0})|*.{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
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
                    foreach (DataRow pRow in IDA_SLC_DOC_FILE.OraSelectData.Rows)
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
            Init_PAYMENT_DATE(STD_YYYYMM.EditValue);
        }
         
        private void ILA_SLC_PERSON_MONTH_SelectedRowData(object pSender)
        {
            IGR_SLC_DOC_ITEM_B.SetCellValue("PAY_SUPPLY_DATE", PAY_SUPPLY_DATE.EditValue);
        }

        private void ILA_SLC_DOC_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SLC_DOC_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_SLC_REASON_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SLC_REASON_CODE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_SLC_INCOME_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SLC_INCOME_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_SLC_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SLC_BANK");
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
            if (iConv.ISNull(e.Row["SLC_DOC_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SLC_DOC_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        }
        
        #endregion
         
    }
}