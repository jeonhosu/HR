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

namespace HRMF0370
{
    public partial class HRMF0370 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0370()
        {
            InitializeComponent();
        }

        public HRMF0370(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultValues()
        {
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            // LEAVE CLOSE TYPE SETTING
            ILD_CLOSED_FLAG_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ILD_CLOSED_FLAG_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            IDC_DEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            IDC_DEFAULT_VALUE.ExecuteNonQuery();
            W_CLOSED_FLAG_NAME.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            W_CLOSED_FLAG.EditValue = IDC_DEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            W_CORP_NAME.BringToFront();
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_WEEK_DATE_FR.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_FR.Focus();
                return;
            }
            if (W_WEEK_DATE_TO.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_TO.Focus();
                return;
            }

            IDA_EX_PERSON.Fill();
            IGR_EX_PERSON.Focus();
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                { 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_EX_PERSON.Cancel();
                    IDA_EX_PERSON_DTL.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if(IDA_EX_PERSON.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Delete Qeustion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }

                        IDC_DELETE_EX_PERSON.SetCommandParamValue("W_WEEK_DATE_FR", IGR_EX_PERSON.GetCellValue("WEEK_DATE_FR"));
                        IDC_DELETE_EX_PERSON.SetCommandParamValue("W_WEEK_DATE_TO", IGR_EX_PERSON.GetCellValue("WEEK_DATE_TO"));
                        IDC_DELETE_EX_PERSON.SetCommandParamValue("W_PERSON_ID", IGR_EX_PERSON.GetCellValue("PERSON_ID"));
                        IDC_DELETE_EX_PERSON.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_EX_PERSON.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_EX_PERSON.GetCommandParamValue("O_MESSAGE"));
                        if(vSTATUS == "F")
                        {
                            if(vMESSAGE != String.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        Search_DB();
                    }
                    else if(IDA_EX_PERSON_DTL.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Delete Qeustion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        
                        IDC_DELETE_EX_PERSON_DTL.SetCommandParamValue("W_WEEK_DATE", IGR_EX_PERSON_DTL.GetCellValue("WEEK_DATE"));
                        IDC_DELETE_EX_PERSON_DTL.SetCommandParamValue("W_PERSON_ID", IGR_EX_PERSON_DTL.GetCellValue("PERSON_ID"));
                        IDC_DELETE_EX_PERSON_DTL.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_EX_PERSON_DTL.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_EX_PERSON_DTL.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != String.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        IDA_EX_PERSON_DTL.Fill(); 
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(IDA_EX_PERSON.IsFocused)
                    {
                        ExcelExport(IGR_EX_PERSON);
                    }
                    else if(IDA_EX_PERSON_DTL.IsFocused)
                    {
                        ExcelExport(IGR_EX_PERSON_DTL);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0370_Load(object sender, EventArgs e)
        {
            DefaultValues();

            W_DUTY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            IDA_EX_PERSON.FillSchema();
        }

        private void HRMF0370_Shown(object sender, EventArgs e)
        {
            
        }

        private void BTN_CAL_EX_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_WEEK_DATE_FR.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_FR.Focus();
                return;
            }
            if (W_WEEK_DATE_TO.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_TO.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_EX_PERSON.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_EX_PERSON.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_EX_PERSON.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if(vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
                return;
            }
            Search_DB();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_WEEK_DATE_FR.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_FR.Focus();
                return;
            }
            if (W_WEEK_DATE_TO.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_TO.Focus();
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10383"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_EX_PERSON_CLOSED.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_EX_PERSON_CLOSED.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_EX_PERSON_CLOSED.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            Search_DB();
        }

        private void BTN_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_WEEK_DATE_FR.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_FR.Focus();
                return;
            }
            if (W_WEEK_DATE_TO.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WEEK_DATE_TO.Focus();
                return;
            }

            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10384"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_EX_PERSON_CLOSED_CANCEL.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_EX_PERSON_CLOSED_CANCEL.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_EX_PERSON_CLOSED_CANCEL.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (vSTATUS == "F")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            Search_DB();
        }

        #endregion

        #region ----- Lookup event ------

        private void ilaYYYYMM_0_SelectedRowData(object pSender)
        {
            W_WEEK_CODE.EditValue = null;
            W_WEEK_DATE_FR.EditValue = null;
            W_WEEK_DATE_TO.EditValue = null;
        }

        private void ILA_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }
         
        private void ILA_WORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("WORK_TYPE", "Y");
        }
         
        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_JOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }
         
        private void ILA_YYYYMM_WEEK_SelectedRowData(object pSender)
        {
            idcYYYYMM_WEEK.SetCommandParamValue("W_WEEK_CODE", W_WEEK_CODE.EditValue);
            idcYYYYMM_WEEK.ExecuteNonQuery();
        }
        #endregion

        #region ------ Adpater event ------

        private void IDA_DAY_LEAVE_WEEK_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        #endregion

    }
}