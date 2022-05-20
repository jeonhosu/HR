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

namespace HRMF0250
{
    public partial class HRMF0250 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0250()
        {
            InitializeComponent();
        }

        public HRMF0250(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public HRMF0250(Form pMainForm, ISAppInterface pAppInterface, object pJOB_NO)
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
            IDA_CERTIFICATE_REQ.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ.Refillable = true;
            IGR_APPROVED_CERTI.LastConfirmChanges();

            IDA_CERTIFICATE_REQ.SetSelectParamValue("W_SOB_ID", -1);
            IDA_CERTIFICATE_REQ.Fill();

            CHECK.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            IDA_CERTIFICATE_REQ.Cancel();
            IDA_CERTIFICATE_REQ.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
            IDA_CERTIFICATE_REQ.Fill();
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
                    IDA_CERTIFICATE_REQ.AddOver();
                    Default_Setting(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_CERTIFICATE_REQ.AddUnder();
                    Default_Setting();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_CERTIFICATE_REQ.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_CERTIFICATE_REQ.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    IDA_CERTIFICATE_REQ.Delete();
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
        
        private void HRMF0250_Load(object sender, EventArgs e)
        {
            IDA_CERTIFICATE_REQ.FillSchema();
        }

        private void HRMF0250_Shown(object sender, EventArgs e)
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
            IDA_CERTIFICATE_REQ.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ.Refillable = true;
        }

        #endregion

        private void Default_Setting()
        {
            IGR_APPROVED_CERTI.SetCellValue("SELECT_YN", "Y");
            IGR_APPROVED_CERTI.SetCellValue("REQ_DATE", DateTime.Today ); 
            IGR_APPROVED_CERTI.SetCellValue("PRINT_COUNT", 1);
            IGR_APPROVED_CERTI.CurrentCellMoveTo(IGR_APPROVED_CERTI.GetColumnToIndex("REQ_DATE"));
            IGR_APPROVED_CERTI.CurrentCellActivate(IGR_APPROVED_CERTI.GetColumnToIndex("REQ_DATE"));
            IGR_APPROVED_CERTI.Focus();
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            int mIDX_SELECT_YN = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");
            ////int mIDX_NAME = IGR_APPROVED_CERTI.GetColumnToIndex("NAME");
            ////int mIDX_CERT_TYPE = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_TYPE_NAME");
            ////int mIDX_REMARK = IGR_APPROVED_CERTI.GetColumnToIndex("REMARK");
            ////int mIDX_SEND_ORG = IGR_APPROVED_CERTI.GetColumnToIndex("SEND_ORG");
            ////int mIDX_DESCRIPTION = IGR_APPROVED_CERTI.GetColumnToIndex("DESCRIPTION");

            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "A")
            {
                BTN_REQ_OK.Enabled = false;
                BTN_REQ_CANCEL.Enabled = false;

                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
                //    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 0;
                //    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 0;
                //    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Updatable = 0;
                //    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Updatable = 0;
                //    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
            }
            else
            {
                if (mAPPROVE_STATE == "N")
                {
                    BTN_REQ_OK.Enabled = true;
                    BTN_REQ_CANCEL.Enabled = false;

                    //IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 1;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 1;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Updatable = 1;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Updatable = 1;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 1;
                }
                else
                {
                    BTN_REQ_OK.Enabled = false;
                    BTN_REQ_CANCEL.Enabled = true;

                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 0;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 0;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Updatable = 0;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Updatable = 0;
                    //        IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
                }
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
            SEARCH_DB();
        }

        private void Init_Grid(int pIDX_Row)
        {
            int mIDX_APPROVE_STATE = IGR_APPROVED_CERTI.GetColumnToIndex("APPROVE_STATE");
            int mIDX_SELECT_YN = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");
            int mIDX_NAME = IGR_APPROVED_CERTI.GetColumnToIndex("NAME");
            int mIDX_CERT_TYPE = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_TYPE");
            int mIDX_CERT_TYPE_NAME = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_TYPE_NAME");
            int mIDX_TASK_DESC = IGR_APPROVED_CERTI.GetColumnToIndex("TASK_DESC");
            int mIDX_REMARK = IGR_APPROVED_CERTI.GetColumnToIndex("REMARK");
            int mIDX_SEND_ORG = IGR_APPROVED_CERTI.GetColumnToIndex("SEND_ORG");
            int mIDX_DESCRIPTION = IGR_APPROVED_CERTI.GetColumnToIndex("DESCRIPTION"); 
            int mIDX_YEAR_YYYY = IGR_APPROVED_CERTI.GetColumnToIndex("YEAR_YYYY");
            int mIDX_MONTH_FR = IGR_APPROVED_CERTI.GetColumnToIndex("MONTH_FR");
            int mIDX_MONTH_TO = IGR_APPROVED_CERTI.GetColumnToIndex("MONTH_TO");
            int mIDX_SAVING_INFO_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("SAVING_INFO_FLAG");
            int mIDX_HOUSE_LEASE_INFO_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("HOUSE_LEASE_INFO_FLAG");

            string mAPPROVE_STATE = iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(pIDX_Row, mIDX_APPROVE_STATE), "N");
            if (mAPPROVE_STATE == "N" || mAPPROVE_STATE == "R")
            {
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE_NAME].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_TASK_DESC].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 1;
                 
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE_NAME].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_TASK_DESC].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Insertable = 1;

                if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(pIDX_Row, mIDX_CERT_TYPE), "N").StartsWith("2"))
                {
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Updatable = 1;

                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Insertable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Insertable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Insertable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Insertable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Insertable = 1;
                } 
                else
                {
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Updatable = 0;

                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Insertable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Insertable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Insertable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Insertable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Insertable = 0;
                }
            }
            else
            {
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE_NAME].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_TASK_DESC].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Updatable = 0;

                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE_NAME].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_TASK_DESC].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_REMARK].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SEND_ORG].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Insertable = 0;
            } 
             
            IGR_APPROVED_CERTI.ResetDraw = true;
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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");
            int vIDX_PRINT_REQ_NUM = IGR_APPROVED_CERTI.GetColumnToIndex("PRINT_REQ_NUM");
            string vSTATUS = "F";
            string vMESSAGE = null;

            IDA_CERTIFICATE_REQ.OraSelectData.AcceptChanges();
            IDA_CERTIFICATE_REQ.Refillable = true;
            IGR_APPROVED_CERTI.LastConfirmChanges();
             
            for (int i = 0; i < IGR_APPROVED_CERTI.RowCount; i++)
            {
                if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(i, vIDX_SELECT_FLAG), "N") == "Y")
                {
                    string vPRINT_REQ_NUM = iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(i, vIDX_PRINT_REQ_NUM));

                    if (!string.IsNullOrEmpty(vPRINT_REQ_NUM))
                    { 
                        IDC_SET_UPDATE_REQUEST.SetCommandParamValue("W_APPROVE_STATUS", pApproved_Flag);
                        IDC_SET_UPDATE_REQUEST.SetCommandParamValue("W_PRINT_REQ_NUM", vPRINT_REQ_NUM);
                        IDC_SET_UPDATE_REQUEST.ExecuteNonQuery();
                        vSTATUS = iConv.ISNull(IDC_SET_UPDATE_REQUEST.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_REQUEST.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_SET_UPDATE_REQUEST.ExcuteError)
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                            Application.DoEvents();
                            MessageBoxAdv.Show(IDC_SET_UPDATE_REQUEST.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if(vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                    }
                }
            }
             
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
        }

        #region ----- Lookup Event ------

        private void ILA_SEARCH_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "CERT_SEARCH_TYPE");
            ILD_SEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ibt_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CERTIFICATE_REQ.Update();

            Set_Update_Approve("A");
        }

        private void ibt_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            Set_Update_Approve("N");
        }

        private void ILA_CERT_REMARK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "CERT_SEND");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CERTIFICATE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_CERTIFICATE_SelectedRowData(object pSender)
        {
            int mIDX_YEAR_YYYY = IGR_APPROVED_CERTI.GetColumnToIndex("YEAR_YYYY");
            int mIDX_MONTH_FR = IGR_APPROVED_CERTI.GetColumnToIndex("MONTH_FR");
            int mIDX_MONTH_TO = IGR_APPROVED_CERTI.GetColumnToIndex("MONTH_TO");
            int mIDX_SAVING_INFO_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("SAVING_INFO_FLAG");
            int mIDX_HOUSE_LEASE_INFO_FLAG = IGR_APPROVED_CERTI.GetColumnToIndex("HOUSE_LEASE_INFO_FLAG");

            if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue("CERT_TYPE"), "N").StartsWith("2"))
            {
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Updatable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Updatable = 1;

                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Insertable = 1;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Insertable = 1;
            }
            else
            {
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Updatable = 0;

                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_YEAR_YYYY].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_FR].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_MONTH_TO].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SAVING_INFO_FLAG].Insertable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_HOUSE_LEASE_INFO_FLAG].Insertable = 0;
            }
        }

        #endregion

    }
}