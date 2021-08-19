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

namespace HRMF0551
{
    public partial class HRMF0551 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #region ----- Variables -----

        


        #endregion;

        #region ----- Constructor -----

        public HRMF0551()
        {
            InitializeComponent();
        }

        public HRMF0551(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

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
            IDA_PAYMENT_TRANSFER.Fill();
            IGR_PAYMENT_TRANSFER.Focus();
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

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xls";

            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, vSaveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = vSaveFileDialog.FileName;
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
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{

                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{
                    //    idaWAGE_TRANSFER_INFO.AddOver();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{
                    //    idaWAGE_TRANSFER_INFO.AddUnder();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{
                    //    idaWAGE_TRANSFER_INFO.Update();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{
                    //    idaWAGE_TRANSFER_INFO.Cancel();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaWAGE_TRANSFER_INFO.IsFocused)
                    //{
                    //    idaWAGE_TRANSFER_INFO.Delete();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_PAYMENT_TRANSFER);
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0551_Load(object sender, EventArgs e)
        {
            CORP_NAME_0.BringToFront();
        }

        private void HRMF0551_Shown(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today); //년월
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);

            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDV_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDV_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDV_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDV_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDV_CORP.GetCommandParamValue("O_CORP_ID");
            CORP_NAME_0.BringToFront();

            IDC_BASE_CURRENCY.ExecuteNonQuery();
            V_CURRENCY_CODE.EditValue = IDC_BASE_CURRENCY.GetCommandParamValue("O_CURRENCY_CODE");
        }

        #endregion

        #region ----- Lookup event -----

        private void ilaYEAR_MONTH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYEAR_MONTH.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYEAR_MONTH.SetLookupParamValue("W_WORK_TERM_TYPE", "D2");
        }

        private void ilaDEPT_ID_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_CORP_ID", CORP_ID_0.EditValue);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPAY_GRADE_ID_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");  // post로 대체해서 적용
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}