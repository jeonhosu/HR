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
using System.IO;

using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0503
{
    public partial class HRMF0503 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----

        public HRMF0503(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            itpDEDUCTION.TabVisible = false;
            if (iConv.ISNull(CORP_TYPE.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();
                igbCORP_GROUP_0.BringToFront();

                irb_ALL_0.RadioButtonValue = "A";
                CORP_TYPE.EditValue = "A";
                CORP_TYPE.BringToFront();
                itpDEDUCTION.TabVisible = true;
            }
            else if (iConv.ISNull(CORP_TYPE.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
                itpDEDUCTION.TabVisible = true;
            }
            
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
            if (CORP_ID_0.EditValue == null&& CORP_TYPE.EditValue.ToString() !="4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (iConv.ISNull(PAY_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Year Month(급여 년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (iConv.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private void Add_Allowance_Insert()
        {// 지급항목.
            igrADD_ALLOWANCE.SetCellValue("PAY_YYYYMM", PAY_YYYYMM_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("WAGE_TYPE", WAGE_TYPE_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("WAGE_TYPE_NAME", WAGE_TYPE_NAME_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("ALLOWANCE_NAME", ALLOWANCE_NAME_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("ADD_ALLOWANCE_ID", ALLOWANCE_ID_0.EditValue);
            igrADD_ALLOWANCE.SetCellValue("TAX_YN", "Y");
            igrADD_ALLOWANCE.SetCellValue("HIRE_INSUR_YN", "Y");
            igrADD_ALLOWANCE.SetCellValue("CREATED_FLAG", "M");
        }

        private void Add_Deduction_Insert()
        {// 공제항목.
            igrADD_DEDUCTION.SetCellValue("PAY_YYYYMM", PAY_YYYYMM_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("WAGE_TYPE", WAGE_TYPE_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("WAGE_TYPE_NAME", WAGE_TYPE_NAME_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("DEDUCTION_NAME", ALLOWANCE_NAME_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("ADD_DEDUCTION_ID", ALLOWANCE_ID_0.EditValue);
            igrADD_DEDUCTION.SetCellValue("CREATED_FLAG", "M");
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null&& CORP_TYPE.EditValue.ToString() != "4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            if (iConv.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }

            if (iConv.ISNull(WAGE_TYPE_0.EditValue) == String.Empty)
            {// 급상여구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            if (TB_PAYMENT_ADDITION.SelectedTab.TabIndex == 1)
            {
                idaADD_ALLOWANCE.Fill();
                igrADD_ALLOWANCE.Focus();
            }
            else if (TB_PAYMENT_ADDITION.SelectedTab.TabIndex == 2)
            {
                idaADD_DEDUCTION.Fill();
                igrADD_DEDUCTION.Focus();
            }            
        }

        private void Show_Import(bool pView_Flag, string pAllowance_Type)
        {
            if (pView_Flag == true)
            {
                if (iConv.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
                {// 급여년월
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    PAY_YYYYMM_0.Focus();
                    return;
                }

                if (iConv.ISNull(WAGE_TYPE_0.EditValue) == String.Empty)
                {// 급상여구분
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WAGE_TYPE_NAME_0.Focus();
                    return;
                }

                ALLOWANCE_TYPE.EditValue = pAllowance_Type;
                PAY_YYYYMM.EditValue = PAY_YYYYMM_0.EditValue;
                WAGE_TYPE_NAME.EditValue = WAGE_TYPE_NAME_0.EditValue;
                WAGE_TYPE.EditValue = WAGE_TYPE_0.EditValue;

                UPLOAD_FILE_PATH.EditValue = String.Empty;
                V_START_ROW.EditValue = 2;
                V_MESSAGE.PromptText = "";
                V_PB_INTERFACE.BarFillPercent = 0;

                igbCORP_GROUP_0.Enabled = false;
                TB_PAYMENT_ADDITION.Enabled = false;

                Application.DoEvents();

                GB_IMPORT.BringToFront();
                GB_IMPORT.Visible = true;
            }
            else
            {
                igbCORP_GROUP_0.Enabled = true;
                TB_PAYMENT_ADDITION.Enabled = true;

                GB_IMPORT.Visible = false;
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

        #region ----- Excel Upload -----

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog1.DefaultExt = "xlsx";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private bool Excel_Upload()
        {
            bool vResult = false;

            if (iConv.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(V_START_ROW.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_START_ROW))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(PAY_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(WAGE_TYPE.EditValue) == String.Empty)
            {// 급상여구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WAGE_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool vXL_Load_OK = false;
            string vOPenFileName = UPLOAD_FILE_PATH.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            V_MESSAGE.PromptText = "Importing Start....";
            try
            {
                if (vXL_Load_OK == true)
                {
                    if (iConv.ISNull(ALLOWANCE_TYPE.EditValue) == "ALLOWANCE")
                    {
                        vXL_Load_OK = vXL_Upload.LoadXL_Allowance(IDC_A_IMPORT_EXCEL, iConv.ISNumtoZero(V_START_ROW.EditValue, 2), V_PB_INTERFACE, V_MESSAGE);
                    }
                    else
                    {  
                        vXL_Load_OK = vXL_Upload.LoadXL_Deduction(IDC_D_IMPORT_EXCEL, iConv.ISNumtoZero(V_START_ROW.EditValue, 2), V_PB_INTERFACE, V_MESSAGE);
                    }
                    if (vXL_Load_OK == false)
                    {
                        vResult = false;
                    }
                    else
                    {
                        V_MESSAGE.PromptText = "Importing Completed....";
                        vResult = true;
                    }
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                vResult = false;
                return vResult;
            }
            vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return vResult;
        }

        #endregion


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
                    if (AddAllowance_Check() == false)
                    {
                        return;
                    }
                    if (idaADD_ALLOWANCE.IsFocused)
                    {                        
                        idaADD_ALLOWANCE.AddOver();
                        Add_Allowance_Insert();
                    }
                    else if (idaADD_DEDUCTION.IsFocused)
                    {
                        idaADD_DEDUCTION.AddOver();
                        Add_Deduction_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (AddAllowance_Check() == false)
                    {
                        return;
                    }
                    if (idaADD_ALLOWANCE.IsFocused)
                    {
                        idaADD_ALLOWANCE.AddUnder();
                        Add_Allowance_Insert();
                    }
                    else if (idaADD_DEDUCTION.IsFocused)
                    {
                        idaADD_DEDUCTION.AddUnder();
                        Add_Deduction_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaADD_ALLOWANCE.IsFocused)
                    {
                        idaADD_ALLOWANCE.Update();
                    }
                    else if (idaADD_DEDUCTION.IsFocused)
                    {
                        idaADD_DEDUCTION.Update();
                    }
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

        private void HRMF0503_Load(object sender, EventArgs e)
        {
            idaADD_ALLOWANCE.FillSchema();
            idaADD_DEDUCTION.FillSchema(); 

            DefaultCorporation();              //Default Corp.
            DefaultPaymentTerm();       //Default Term.     
            Show_Import(false, "");
        }

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void itbPAYMENT_ADDITION_Click(object sender, EventArgs e)
        {
            if (TB_PAYMENT_ADDITION.SelectedTab.TabIndex == 1)
            {
                ALLOWANCE_NAME_0.EditValue = null;
                ALLOWANCE_ID_0.EditValue = null;
                igrADD_ALLOWANCE.Focus();
            }
            else if (TB_PAYMENT_ADDITION.SelectedTab.TabIndex == 2)
            {
                ALLOWANCE_NAME_0.EditValue = null;
                ALLOWANCE_ID_0.EditValue = null;
                igrADD_DEDUCTION.Focus();
            }
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
            if (iConv.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Start Year Month(시작년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Item(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void idaADD_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
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
            if (iConv.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Start Year Month(시작년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Wage Type(급상여 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DEDUCTION_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Deduction Item(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DEDUCTION_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Deduction Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
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

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }
        private void ilaALLOWANCE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "1 = 1 ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

            ildCOMMON_GROUP.SetLookupParamValue("W_GROUP_CODE", ALLOWANCE_TYPE_0.EditValue);
            ildCOMMON_GROUP.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'Y'");
            ildCOMMON_GROUP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", " HC.VALUE1 = 'Y' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEDUCTION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "DEDUCTION");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'Y'");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Excel UpLoad -----

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }
        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Excel_Upload() == true)
            {
                Show_Import(false, "");
                Search_DB();
            }
        }

        private void BTN_CLOSED_I_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Import(false, "");
        }
         
        private void bXL_UpLoad_Allowance_ButtonClick(object pSender, EventArgs pEventArgs)
        {

            Show_Import(true, "ALLOWANCE"); 
        }
         
        private void bXL_UpLoad_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Import(true, "DEDUCTION"); 
        }

        #endregion
         

    }
}