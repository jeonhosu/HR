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

using System.IO;
using Syncfusion.GridExcelConverter;

namespace HRMF0604
{
    public partial class HRMF0604 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0604(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
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
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_RESERVE_YYYYMM.EditValue == null)
            {// 정산년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SDM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            int vTP_INDEX = TB_MAIN.SelectedTab.TabIndex;
            string vPERSON_NUM = string.Empty;
            if (TP_DETAIL.TabIndex == vTP_INDEX)
            {
                vPERSON_NUM = iString.ISNull(IGR_RETIRE_RESERVE.GetCellValue("PERSON_NUM"));
                IDA_RETIRE_RESERVE.Fill();
                int vIDX_PERSON_NUM = IGR_RETIRE_RESERVE.GetColumnToIndex("PERSON_NUM");
                for (int r = 0; r < IGR_RETIRE_RESERVE.RowCount; r++)
                {
                    if (vPERSON_NUM == iString.ISNull(IGR_RETIRE_RESERVE.GetCellValue(r, vIDX_PERSON_NUM)))
                    {
                        IGR_RETIRE_RESERVE.CurrentCellMoveTo(vIDX_PERSON_NUM);
                        IGR_RETIRE_RESERVE.CurrentCellActivate(vIDX_PERSON_NUM);
                        return;
                    }
                }
                IGR_RETIRE_RESERVE.Focus();
            }
            else if (TP_SUM.TabIndex == vTP_INDEX)
            {
                vPERSON_NUM = iString.ISNull(IGR_RETIRE_RESERVE_SUM.GetCellValue("PERSON_NUM"));
                IDA_RETIRE_RESERVE_SUM.Fill();
                int vIDX_PERSON_NUM = IGR_RETIRE_RESERVE_SUM.GetColumnToIndex("PERSON_NUM");
                for (int r = 0; r < IGR_RETIRE_RESERVE_SUM.RowCount; r++)
                {
                    if (vPERSON_NUM == iString.ISNull(IGR_RETIRE_RESERVE_SUM.GetCellValue(r, vIDX_PERSON_NUM)))
                    {
                        IGR_RETIRE_RESERVE_SUM.CurrentCellMoveTo(vIDX_PERSON_NUM);
                        IGR_RETIRE_RESERVE_SUM.CurrentCellActivate(vIDX_PERSON_NUM);
                        return;
                    }
                }
                IGR_RETIRE_RESERVE_SUM.Focus();
            } 
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
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                        if (TB_MAIN.SelectedTab.TabIndex == TP_DETAIL.TabIndex)
                        {
                            ExcelExport(IGR_RETIRE_RESERVE);
                        }
                        else if (TB_MAIN.SelectedTab.TabIndex == TP_SUM.TabIndex)
                        {
                            ExcelExport(IGR_RETIRE_RESERVE_SUM);
                        }

                        
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
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";
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
                                    Syncfusion.GridExcelConverter.ConverterOptions.RowHeaders);  // 엑셀 다운로드 뒤 라인수 안맞을때 바꿔서 써봐야함   ColumnHeaders
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


        #region ----- Form Event -----

        private void HRMF0604_Load(object sender, EventArgs e)
        {
            this.Visible = true;

            // Year Month Setting
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
            W_RESERVE_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);

            // CORP SETTING
            DefaultCorporation();           

            //퇴직연금 기본값 조회 
            IDC_DEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "RETIRE_IRP_TYPE");
            IDC_DEFAULT_VALUE_GROUP.ExecuteNonQuery();
            string vCODE = iString.ISNull(IDC_DEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE"));
            if (vCODE == "DC")
            {
                BTN_SET_PAYMENT.Visible = true;
            }
            else
            {
                BTN_SET_PAYMENT.Visible = false;
            }
         
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            IDA_RETIRE_RESERVE.FillSchema();            
        }

        private void BTN_SET_RETIRE_RESERVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == string.Empty)
            {// 정산년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SDM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            IDC_SET_RETIRE_RESERVE.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_SET_RETIRE_RESERVE.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_SET_RETIRE_RESERVE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_RETIRE_RESERVE.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Search_DB();
        }

        private void BTN_SET_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == string.Empty)
            {// 정산년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SDM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            HRMF0604_PAYMENT vHRMF0604_PAYMENT = new HRMF0604_PAYMENT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                    , W_RESERVE_YYYYMM.EditValue 
                                                                    , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                    , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                    , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue 
                                                                    , W_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0604_PAYMENT.ShowDialog();
            vHRMF0604_PAYMENT.Dispose();
        }
        
        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == string.Empty)
            {// 정산년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SDM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            HRMF0604_CLOSED vHRMF0604_CLOSED = new HRMF0604_CLOSED(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                    , W_RESERVE_YYYYMM.EditValue
                                                                    , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                    , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                    , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue
                                                                    , W_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0604_CLOSED.ShowDialog();
            vHRMF0604_CLOSED.Dispose();
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaMONTH_TOTAL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void IDA_RETIRE_RESERVE_UpdateCompleted(object pSender)
        {
            Search_DB();
        }

        #endregion

        #region ----- LookUp Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_CORP_TYPE", "1");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion



    }
}