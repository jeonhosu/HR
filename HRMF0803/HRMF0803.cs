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

using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace HRMF0803
{
    public partial class HRMF0803 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;
        
        #region ----- Constructor -----

        public HRMF0803(Form pMainForm, ISAppInterface pAppInterface)
        {
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

        private void DefaultCorp()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_USABLE_CHECK_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {
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
            if (TB_MAIN.SelectedTab.TabIndex == TP_FOOD_TIME.TabIndex)
            {
                idaPERSON_SUMMARY.Fill();

                igrFOOD_SUMMARY.Focus();
            }
            else
            {
                IDA_FOOD_DED.Fill();
                IGR_FOOD_DED.Focus();
            }
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

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
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
                    IDA_FOOD_DED.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_FOOD_DED.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(TB_MAIN.SelectedTab.TabIndex == TP_FOOD_DED.TabIndex)
                    {
                        ExcelExport(IGR_FOOD_DED);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0803_Load(object sender, EventArgs e)
        {
            W_FOOD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISGetDate();

            V_RB_CLOSED_YES.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_STATUS.EditValue = V_RB_CLOSED_YES.RadioCheckedString;

            DefaultCorp();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]    
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if(V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS.EditValue = V_RB_ALL.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_NO_Click(object sender, EventArgs e)
        {
            if (V_RB_CLOSED_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS.EditValue = V_RB_CLOSED_NO.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_YES_Click(object sender, EventArgs e)
        {
            if (V_RB_CLOSED_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS.EditValue = V_RB_CLOSED_YES.RadioCheckedString;
            }
        }

        private void V_RB_PAYMENT_Click(object sender, EventArgs e)
        {
            if (V_RB_PAYMENT.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS.EditValue = V_RB_PAYMENT.RadioCheckedString;
            }
        }

        private void BTN_FOOD_DED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
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
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_FOOD_DED.ExecuteNonQuery();
            string O_STATUS = iConv.ISNull(IDC_EXEC_FOOD_DED.GetCommandParamValue("O_STATUS"));
            string O_MESSAGE = iConv.ISNull(IDC_EXEC_FOOD_DED.GetCommandParamValue("O_MESSAGE"));
            if(O_STATUS == "F")
            {
                if(O_MESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(O_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
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
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_FOOD_DED_CLOSED.SetCommandParamValue("W_CLOSED_STATUS", "OK");
            IDC_EXEC_FOOD_DED_CLOSED.ExecuteNonQuery();
            string O_STATUS = iConv.ISNull(IDC_EXEC_FOOD_DED_CLOSED.GetCommandParamValue("O_STATUS"));
            string O_MESSAGE = iConv.ISNull(IDC_EXEC_FOOD_DED_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (O_STATUS == "F")
            {
                if (O_MESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(O_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
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
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_EXEC_FOOD_DED_CLOSED.SetCommandParamValue("W_CLOSED_STATUS", "CANCEL");
            IDC_EXEC_FOOD_DED_CLOSED.ExecuteNonQuery();
            string O_STATUS = iConv.ISNull(IDC_EXEC_FOOD_DED_CLOSED.GetCommandParamValue("O_STATUS"));
            string O_MESSAGE = iConv.ISNull(IDC_EXEC_FOOD_DED_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (O_STATUS == "F")
            {
                if (O_MESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(O_MESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void BTN_TRANS_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
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

            HRMF0803_SUB vHRMF0803_SUB = new HRMF0803_SUB(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                        , W_FOOD_YYYYMM.EditValue 
                                                        , W_START_DATE.EditValue, W_END_DATE.EditValue
                                                        , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                        , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue 
                                                        , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue
                                                        , "OK");
            vHRMF0803_SUB.ShowDialog();
            vHRMF0803_SUB.Dispose();
        }

        private void BTN_TRANS_PAYMENT_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
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

            HRMF0803_SUB vHRMF0803_SUB = new HRMF0803_SUB(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                        , W_FOOD_YYYYMM.EditValue
                                                        , W_START_DATE.EditValue, W_END_DATE.EditValue
                                                        , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                        , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                        , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue
                                                        , "CANCEL");
            vHRMF0803_SUB.ShowDialog();
            vHRMF0803_SUB.Dispose();
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_FOOD_DED_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if(pBindingManager.DataRow == null)
            {
                return;
            }

            int vIDX_FOOD_DED_FLAG = IGR_FOOD_DED.GetColumnToIndex("FOOD_DED_FLAG");
            IGR_FOOD_DED.GridAdvExColElement[vIDX_FOOD_DED_FLAG].Updatable = pBindingManager.DataRow["UPDATABLE_FLAG"]; 
        }

        #endregion

        #region ----- LookUp Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 2)));
        }

        private void ILA_FOOD_FLAG_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FOOD_FLAG");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}