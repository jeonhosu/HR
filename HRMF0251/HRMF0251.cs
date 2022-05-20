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

namespace HRMF0251
{
    public partial class HRMF0251 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #endregion;

        #region ----- Constructor -----

        public HRMF0251()
        {
            InitializeComponent();
        }

        public HRMF0251(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public HRMF0251(Form pMainForm, ISAppInterface pAppInterface, object pJOB_NO)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_CERTI_TYPE_NAME.EditValue = pJOB_NO;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            IDA_CERTIFICATE_REQ_APPR.SetSelectParamValue("W_SOB_ID", -1);
            IDA_CERTIFICATE_REQ_APPR.Fill();

            CHECK.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
             
            IDA_CERTIFICATE_REQ_APPR.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID); 
            IDA_CERTIFICATE_REQ_APPR.Cancel();
            IDA_CERTIFICATE_REQ_APPR.Fill();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
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
                    IDA_CERTIFICATE_REQ_APPR.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //IDA_APPROVED_CERTI.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //ExcelExport(IGR_APPROVED_CERTI);
                }
            }
        }

        #endregion;

        #region ----- Excel Export -----
        private void ExcelExport(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ////데이터 테이블을 이용한 export
                //Syncfusion.XlsIO.ExcelEngine vEng = new Syncfusion.XlsIO.ExcelEngine();
                //Syncfusion.XlsIO.IApplication vApp = vEng.Excel;
                //string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();
                //if (vFileExtension == "XLSX")
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel2007;
                //}
                //else
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel97to2003;
                //}
                //Syncfusion.XlsIO.IWorkbook vWorkbook = vApp.Workbooks.Create(1);
                //Syncfusion.XlsIO.IWorksheet vSheet = vWorkbook.Worksheets[0];
                //foreach(System.Data.DataRow vRow in IDA_MATERIAL_LIST_ALL.CurrentRows)
                //{
                //    vSheet.ImportDataTable(vRow.Table, true, 1, 1, -1, -1);
                //}
                //vWorkbook.SaveAs(saveFileDialog.FileName);
                vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);
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
        
        private void HRMF0251_Load(object sender, EventArgs e)
        {
            IDA_CERTIFICATE_REQ_APPR.FillSchema();
        }

        private void HRMF0251_Shown(object sender, EventArgs e)
        {
            //DEFAULT Date SETTING
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);
            
            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            APPROVED_CANCEL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_APPROVE_STATUS.EditValue = "N";

            W_CORP_NAME.BringToFront();

            IDC_CERT_SEARCH_TYPE.SetCommandParamValue("W_GROUP_CODE", "CERT_SEARCH_TYPE");
            IDC_CERT_SEARCH_TYPE.ExecuteNonQuery();
            W_SEARCH_TYPE.EditValue = IDC_CERT_SEARCH_TYPE.GetCommandParamValue("O_CODE");
            W_SEARCH_TYPE_NAME.EditValue = IDC_CERT_SEARCH_TYPE.GetCommandParamValue("O_CODE_NAME");
        }
          
        #region ----- Form Event ------

        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int r = 0; r < IGR_APPROVED_CERTI.RowCount; r++)
            {
                IGR_APPROVED_CERTI.SetCellValue(r, IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN"), CHECK.CheckBoxString);
            }
            IGR_APPROVED_CERTI.LastConfirmChanges();
            IDA_CERTIFICATE_REQ_APPR.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ_APPR.Refillable = true;
        }

        private bool Set_Update_Return(DateTime pSys_Date)
        {
            if (IGR_APPROVED_CERTI.RowCount < 1)
            {
                return false;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_APPROVED_CERTI.LastConfirmChanges();
            IDA_CERTIFICATE_REQ_APPR.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ_APPR.Refillable = true;

            int vIDX_SELECT_YN = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");
            int vIDX_DUTY_PERIOD_ID = IGR_APPROVED_CERTI.GetColumnToIndex("DUTY_PERIOD_ID");
            int vIDX_START_DATE = IGR_APPROVED_CERTI.GetColumnToIndex("START_DATE");
            int vIDX_END_DATE = IGR_APPROVED_CERTI.GetColumnToIndex("END_DATE");
            int vIDX_PERSON_ID = IGR_APPROVED_CERTI.GetColumnToIndex("PERSON_ID");
            int vIDX_APPROVE_STATUS = IGR_APPROVED_CERTI.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_APPROVED_CERTI.RowCount; i++)
            {
                if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_DUTY_PERIOD_ID", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_DUTY_PERIOD_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_CHECK_YN", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_START_DATE", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_START_DATE));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_END_DATE", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_END_DATE));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_PERSON_ID", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_PERSON_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_APPROVE_STATUS", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_SYS_DATE", pSys_Date);
                    IDC_UPDATE_RETURN_TEMP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_RETURN_TEMP.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return false;
                    }
                }
            }
            return true;
        }
        private void ibt_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Update_Approve("C");
        }

        private void ibt_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Update_Approve("A");
        }

        private void ibt_REJECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_START_DATE.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_START_DATE.Focus();
                return;
            }
            if (W_END_DATE.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_END_DATE.Focus();
                return;
            }
            if (Convert.ToDateTime(W_START_DATE.EditValue) > Convert.ToDateTime(W_END_DATE.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_START_DATE.Focus();
                return;
            }

            Application.UseWaitCursor = true; 
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            DialogResult dlgResultValue;
              
            Form vHRMF0251_RETURN = new HRMF0251_RETURN(isAppInterfaceAdv1.AppInterface
                                                        , W_CORP_ID.EditValue
                                                        , iDate.ISGetDate(W_START_DATE.EditValue)
                                                        , iDate.ISGetDate(W_END_DATE.EditValue)
                                                        );
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0251_RETURN, isAppInterfaceAdv1.AppInterface);
            dlgResultValue = vHRMF0251_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            try
            {
                vHRMF0251_RETURN.Dispose();
            }
            catch
            {

            }
            SEARCH_DB();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false; 
        }

        #endregion

        private void Default_Setting()
        {
            IGR_APPROVED_CERTI.SetCellValue("PRINT_DATE", DateTime.Today );
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);

            int mIDX_REJECT_REMARK = IGR_APPROVED_CERTI.GetColumnToIndex("REJECT_REMARK");
            int mIDX_SELECT_YN = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");

            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "A" || mAPPROVE_STATE == "R" )
            {
                BTN_OK.Enabled = false;
                BTN_CANCEL.Enabled = false;
                BTN_REJECT.Enabled = false;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REJECT_REMARK].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                if (mAPPROVE_STATE == "N")
                {
                    BTN_OK.Enabled = true;
                    BTN_CANCEL.Enabled = false;
                    BTN_REJECT.Enabled = true;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REJECT_REMARK].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
                }
                else
                {
                    BTN_OK.Enabled = false;
                    BTN_CANCEL.Enabled = true;
                    BTN_REJECT.Enabled = true;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REJECT_REMARK].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
                }
            }
            SEARCH_DB();
        }
         
        private void APPROVED_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();

        }

        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (IGR_APPROVED_CERTI.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");  
            int vIDX_REJECT_REMARK = IGR_APPROVED_CERTI.GetColumnToIndex("REJECT_REMARK"); 
            int vIDX_PRINT_REQ_NUM = IGR_APPROVED_CERTI.GetColumnToIndex("PRINT_REQ_NUM");
            string vSTATUS = "F";
            string vMESSAGE = null; 

            IDA_CERTIFICATE_REQ_APPR.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ_APPR.Refillable = true;
            IGR_APPROVED_CERTI.LastConfirmChanges();

            if (pApproved_Flag == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                IGR_APPROVED_CERTI.Focus();
                return;
            }

            for (int i = 0; i < IGR_APPROVED_CERTI.RowCount; i++)
            {
                if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(i, vIDX_SELECT_FLAG), "N") == "Y")
                {
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_CORP_ID", W_CORP_ID.EditValue);
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_APPROVE_STATUS", pApproved_Flag);
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_PRINT_REQ_NUM", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_PRINT_REQ_NUM));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_REJECT_REMARK", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_REJECT_REMARK));
                    IDC_SET_UPDATE_APPROVE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            // eMail 전송.
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
        }
        #region ----- Lookup Event ------
         
        private void ILA_SEARCH_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "CERT_SEARCH_TYPE");
            ILD_SEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        private void ilaJOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DEPT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion 
    }
}