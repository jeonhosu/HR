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

namespace HRMF0218
{
    public partial class HRMF0218 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        private ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0218()
        {
            InitializeComponent();
        }

        public HRMF0218(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_CONTRACT_DTL.TabIndex)
            {
                if(iConv.ISNull(W1_CHANGE_DATE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W1_CHANGE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W1_CHANGE_DATE.Focus();
                    return;
                }
                if (iConv.ISNull(W1_DUE_DAY.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W1_CHANGE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W1_DUE_DAY.Focus();
                    return;
                }

                IDA_CHANGE_CONTRACT.Fill();
                IGR_CHANGE_CONTRACT.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_CONTRACT_LIST.TabIndex)
            {
                if (iConv.ISNull(W2_PERIOD_DATE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_PERIOD_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W2_PERIOD_DATE_FR.Focus();
                    return;
                }
                if (iConv.ISNull(W2_PERIOD_DATE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_PERIOD_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W2_PERIOD_DATE_TO.Focus();
                    return;
                }
                if (iConv.ISNull(W2_CONTRACT_DATE_TYPE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W2_CONTRACT_DATE_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W2_CONTRACT_DATE_TYPE_NAME.Focus();
                    return;
                }

                IDA_CONTRACT_LIST.Fill();
                IGR_CONTRACT_LIST.Focus();
            }
        }

        private void Insert_DB()
        {
            IGR_CHANGE_CONTRACT.SetCellValue("NEW_CONTRACT_DATE", W1_CHANGE_DATE.EditValue);
            IGR_CHANGE_CONTRACT.Focus();
        }

        private void Delete_Contract(object pCONTRACT_ID)
        { 
            if(IGR_CONTRACT_HISTORY.RowCount < 0)
            {
                return;
            }
            if (iConv.ISNull(pCONTRACT_ID) == string.Empty)
            {
                return;
            }
            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            IDC_DELETE_CONTRACT.SetCommandParamValue("P_CONTRACT_ID", pCONTRACT_ID);
            IDC_DELETE_CONTRACT.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_DELETE_CONTRACT.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_DELETE_CONTRACT.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            IDA_CONTRACT_HISTORY.Fill();
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
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

                //xls 历厘规过
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 颇老 历厘 规过
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


        #region ----- MDi ToolBar Button Event -----

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
                    IDA_CHANGE_CONTRACT.AddOver();
                    Insert_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_CHANGE_CONTRACT.AddUnder();
                    Insert_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    System.Windows.Forms.SendKeys.Send("{TAB}");
                    IDA_CHANGE_CONTRACT.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {                    
                    if (IDA_CHANGE_CONTRACT.IsFocused)
                    {
                        IDA_CHANGE_CONTRACT.Cancel();
                    }
                    else if (IDA_CONTRACT_HISTORY.IsFocused)
                    {
                        IDA_CONTRACT_HISTORY.Cancel();
                    }
                    else if (IDA_CONTRACT_LIST.IsFocused)
                    {
                        IDA_CONTRACT_LIST.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CONTRACT_HISTORY.IsFocused)
                    {
                        Delete_Contract(IGR_CONTRACT_HISTORY.GetCellValue("CONTRACT_ID"));
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(IDA_CONTRACT_LIST.IsFocused)
                    {
                        ExcelExport(IGR_CONTRACT_LIST);
                    }
                }
            }
        }

        #endregion;

        #region ----- Convert String Method ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- Convert DateTime Methods ----

        private System.DateTime ConvertDateTime(object pObject)
        {
            System.DateTime vDateTime = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        vDateTime = (System.DateTime)pObject;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vDateTime;
        }

        #endregion;

        #region ----- Private Method ----

        private void DefaultCorporation()
        {
            W1_CHANGE_DATE.EditValue = System.DateTime.Today;
            W1_DUE_DAY.EditValue = 30;

            W2_PERIOD_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W2_PERIOD_DATE_TO.EditValue = DateTime.Today;

            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W1_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W1_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W2_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W2_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            IDC_DEFAULT_CONTRACT_DATE_TYPE.ExecuteNonQuery();
            W2_CONTRACT_DATE_TYPE_NAME.EditValue = IDC_DEFAULT_CONTRACT_DATE_TYPE.GetCommandParamValue("O_CONTRACT_DATE_TYPE_NAME");
            W2_CONTRACT_DATE_TYPE.EditValue = IDC_DEFAULT_CONTRACT_DATE_TYPE.GetCommandParamValue("O_CONTRACT_DATE_TYPE"); 

            W1_CORP_NAME.BringToFront();
            W2_CORP_NAME.BringToFront();
            V_PM_HISTORY.BringToFront();            
        }
         
        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }
         
        #endregion;

        #region ----- Form Event -----

        private void HRMF0218_Load(object sender, EventArgs e)
        {
            DefaultCorporation(); 
            IDA_CHANGE_CONTRACT.FillSchema();
        }

        private void IGR_CHANGE_CONTRACT_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_NEW_CONTRACT_DATE_FR = IGR_CHANGE_CONTRACT.GetColumnToIndex("NEW_CONTRACT_DATE_FR");
            if(e.ColIndex == vIDX_NEW_CONTRACT_DATE_FR)
            {
                int vCONTRACT_MONTH = iConv.ISNumtoZero(IGR_CHANGE_CONTRACT.GetCellValue("CONTRACT_MONTH"));
                DateTime vCONTRACT_DATE_FR = iDate.ISGetDate(e.NewValue);
                vCONTRACT_DATE_FR = iDate.ISDate_Add(vCONTRACT_DATE_FR, 1);
                DateTime vCONTRACT_DATE_TO = iDate.ISDate_Month_Add(vCONTRACT_DATE_FR, vCONTRACT_MONTH);
                IGR_CHANGE_CONTRACT.SetCellValue("NEW_CONTRACT_DATE_TO", vCONTRACT_DATE_TO);
            }
        }

        #endregion;

        #region ----- LookUP Event ----

        private void ILA_DEPT_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_CONTRACT_LEVEL_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_LEVEL", "Y"); 
        }

        private void ILA_POST_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("POST", "Y");
        }
         
        private void ILA_FLOOR_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y"); 
        }
         
        private void ILA_CONTRACT_DATE_TYPE_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_DATE_TYPE", "Y");
        }

        private void ILA_DEPT_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_FLOOR_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_POST_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("POST", "Y");
        }

        private void ILA_CONTRACT_LEVEL_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_LEVEL", "Y");
        }

        private void ILA_PERSON_W1_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_DEPT_ID", W1_DEPT_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_POST_ID", W1_POST_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_JOB_CLASS_ID", null);
            ILD_PERSON.SetLookupParamValue("W_FLOOR_ID", W1_FLOOR_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", W1_CHANGE_DATE.EditValue); 
        }

        private void ILA_PERSON_W2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_DEPT_ID", W2_DEPT_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_POST_ID", W2_POST_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_JOB_CLASS_ID", null);
            ILD_PERSON.SetLookupParamValue("W_FLOOR_ID", W2_FLOOR_ID.EditValue);
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", W2_PERIOD_DATE_TO.EditValue);
        }

        private void ILA_CONTRACT_LEVEL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            object vCONTRACT_DATE = IGR_CHANGE_CONTRACT.GetCellValue("CONTRACT_DATE_TO");
            if(iConv.ISNull(vCONTRACT_DATE) == String.Empty)
            {
                vCONTRACT_DATE = IGR_CHANGE_CONTRACT.GetCellValue("JOIN_DATE");
            }
            ILD_CONTRACT_LEVEL.SetLookupParamValue("W_CONTRACT_DATE", vCONTRACT_DATE);
            ILD_CONTRACT_LEVEL.SetLookupParamValue("W_JOB_CLASS_ID", IGR_CHANGE_CONTRACT.GetCellValue("JOB_CLASS_ID"));
            ILD_CONTRACT_LEVEL.SetLookupParamValue("W_CONTRACT_SEQ_NO", IGR_CHANGE_CONTRACT.GetCellValue("CONTRACT_SEQ_NO"));
            ILD_CONTRACT_LEVEL.SetLookupParamValue("W_ENABLED_FLAG", "Y"); 
        }

        #endregion

        #region ----- Adapter Event ----

        private void IDA_CHANGE_WORK_TYPE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }
            if (iConv.ISNull(e.Row["NEW_CONTRACT_DATE"]) == string.Empty)
            {
                object vMESSAGE = Get_Grid_Prompt(IGR_CHANGE_CONTRACT, IGR_CHANGE_CONTRACT.GetColumnToIndex("NEW_CONTRACT_DATE"));
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", vMESSAGE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["NEW_CONTRACT_DATE_FR"]) == string.Empty)
            {
                object vMESSAGE = Get_Grid_Prompt(IGR_CHANGE_CONTRACT, IGR_CHANGE_CONTRACT.GetColumnToIndex("NEW_CONTRACT_DATE_FR"));
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", vMESSAGE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["NEW_CONTRACT_DATE_TO"]) == string.Empty)
            {
                object vMESSAGE = Get_Grid_Prompt(IGR_CHANGE_CONTRACT, IGR_CHANGE_CONTRACT.GetColumnToIndex("NEW_CONTRACT_DATE_TO"));
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", vMESSAGE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void IDA_CHANGE_CONTRACT_UpdateCompleted(object pSender)
        {
            IDA_CONTRACT_HISTORY.Fill();
        }

        #endregion

    }
}