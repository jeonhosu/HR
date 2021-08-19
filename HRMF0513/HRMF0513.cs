using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;


namespace HRMF0513
{
    public partial class HRMF0513 : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public HRMF0513()
        {
            InitializeComponent();
        }

        public HRMF0513(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(STD_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYYMM_0.Focus();
                return;
            }
            idaBANK_ACCOUNT.Fill();
            igrBANK_ACCOUNT.Focus();
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
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaBANK_ACCOUNT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                        idaBANK_ACCOUNT.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (idaBANK_ACCOUNT.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0513_Load(object sender, EventArgs e)
        {
            idaBANK_ACCOUNT.FillSchema();
        }

        private void HRMF0513_Shown(object sender, EventArgs e)
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


            // Standard Date SETTING   
            STD_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
        }

        #endregion

        #region ----- Lookup event -----

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPAY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        
        private void ilaPAY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        
        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BANK");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
            
        #endregion

        #region ----- Excel Upload : Asset Master -----

        private void Select_Excel_File_10()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FILE_PATH_MASTER.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    FILE_PATH_MASTER.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private void Excel_Upload_10()
        {
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            bool vXL_Load_OK = false;

            if (iString.ISNull(FILE_PATH_MASTER.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH_MASTER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vOPenFileName = FILE_PATH_MASTER.EditValue.ToString();
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
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }


            // 업로드 아답터 fill //
            IDA_EXCEL_BANK_ACCOUNT.Fill();

            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL_10(IDA_EXCEL_BANK_ACCOUNT, 2);

                    if (vXL_Load_OK == false)
                    {
                        IDA_EXCEL_BANK_ACCOUNT.Cancel();
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        IDA_EXCEL_BANK_ACCOUNT.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_EXCEL_BANK_ACCOUNT.Cancel();

                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }
            vXL_Upload.DisposeXL();


            //if (IDA_EXCEL_BANK_ACCOUNT.IsUpdateCompleted == true)
            //{
            //    vSTATUS = "F";
            //    vMESSAGE = string.Empty;

            //    IDC_SET_TRANS_MASTER_TEMP.ExecuteNonQuery();
            //    vSTATUS = iString.ISNull(IDC_SET_TRANS_MASTER_TEMP.GetCommandParamValue("O_STATUS"));
            //    vMESSAGE = iString.ISNull(IDC_SET_TRANS_MASTER_TEMP.GetCommandParamValue("O_MESSAGE"));

            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();
            //    if (IDC_SET_TRANS_MASTER_TEMP.ExcuteError || vSTATUS == "F")
            //    {
            //        if (vMESSAGE != string.Empty)
            //        {
            //            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            return;
            //        }
            //    }
            //}

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion


        private void bXL_Choice_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File_10();
        }

        private void bXL_UpLoad_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Excel_Upload_10();  // 계좌 업로드 //   
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }


    }
}