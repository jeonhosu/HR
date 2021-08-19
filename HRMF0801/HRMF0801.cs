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
using Syncfusion.XlsIO;


namespace HRMF0801
{
    public partial class HRMF0801 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;
        
        #region ----- Constructor -----

        public HRMF0801(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
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

        private void Search_DB()
        {
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

            if (TB_MAIN.SelectedTab.TabIndex == TP_SUM.TabIndex)
            {
                IDA_FOOD_SUMMARY.SetSelectParamValue("W_SEARCH_TYPE", "R");
                IDA_FOOD_SUMMARY.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PERSON.TabIndex)
            {
                IDA_FOOD_PERSON.Fill();
                IGR_FOOD_PERSON.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_VISITOR.TabIndex)
            {
                IDA_FOOD_VISITOR.Fill();
                IGR_FOOD_VISITOR.Focus();
            }
        }
        
        private bool isAdd_DB_Check()
        {// 데이터 추가시 검증.
            if (W_FOOD_DEVICE_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_FOOD_DEVICE_NAME.Focus();
                return false;
            }
            return true;
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
                    if(IDA_FOOD_SUMMARY.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_FOOD_SUMMARY.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_FOOD_SUMMARY.IsFocused)
                    {                
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_FOOD_SUMMARY.IsFocused)
                    {
                        IDA_FOOD_SUMMARY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_FOOD_SUMMARY.IsFocused)
                    {
                        //IDA_FOOD_SUMMARY.Delete();
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(IDA_FOOD_SUMMARY.IsFocused)
                    {
                        ExcelExport(igrFOOD_SUMMARY);
                    }
                    else if(IDA_FOOD_DAY_COUNT.IsFocused)
                    {
                        ExcelExport(igrFOOD_DAY_COUNT);
                    }
                    else if(IDA_FOOD_PERSON.IsFocused)
                    {
                        ExcelExport(IGR_FOOD_PERSON);
                    }
                    else if(IDA_FOOD_VISITOR.IsFocused)
                    {
                        ExcelExport(IGR_FOOD_VISITOR);
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0801_Load(object sender, EventArgs e)
        {
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = DateTime.Today;

            V_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_STATUS.EditValue = V_RB_ALL.RadioCheckedString;

            GB_STATUS.BringToFront();
            W_ALL_FLAG.BringToFront();

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]       
            IDA_FOOD_SUMMARY.FillSchema();
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if (V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
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


        private void BTN_SET_FOOD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
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
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
             
            IDC_SET_FOOD.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_FOOD.GetCommandParamValue("O_STATUS"));
            string vMessage = iConv.ISNull(IDC_SET_FOOD.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
            if(vSTATUS == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMessage != string.Empty)
            {
                MessageBoxAdv.Show(vMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
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

        #endregion

        #region ----- Adapter Event -----

        #endregion

        #region ----- LookUp Event ----- 

        private void ILA_CAFETERIA_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CAFETERIA.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        #endregion
    }
}