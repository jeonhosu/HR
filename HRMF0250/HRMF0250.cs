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

            W_CERTI_TYPE.EditValue = pJOB_NO;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            IDA_APPROVED_CERTI.Cancel();
            IDA_APPROVED_CERTI.Fill();
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
                    Default_Setting();
                    IDA_APPROVED_CERTI.AddOver();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    Default_Setting();
                    IDA_APPROVED_CERTI.AddUnder();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_APPROVED_CERTI.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_APPROVED_CERTI.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    IDA_APPROVED_CERTI.Delete();
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
            IDA_APPROVED_CERTI.FillSchema();
        }

        private void HRMF0250_Shown(object sender, EventArgs e)
        {
            //DEFAULT Date SETTING
            iSTART_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            iEND_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);
            
            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            APPROVED_CANCEL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_APPROVE_STATUS.EditValue = "N";

            W_CORP_NAME_0.BringToFront();

        }
        private void IGR_OPERATION_CAPA_Click(object sender, EventArgs e)
        {

        }

        private void IGR_OPERATION_CAPA_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {       

        }

        #region ----- Form Event ------

        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e)
        {

            for (int r = 0; r < IGR_APPROVED_CERTI.RowCount; r++)
            {
                IGR_APPROVED_CERTI.SetCellValue(r, IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN"), CHECK.CheckBoxString);
            }
            IGR_APPROVED_CERTI.LastConfirmChanges();
            IDA_APPROVED_CERTI.OraSelectData.AcceptChanges();
            IDA_APPROVED_CERTI.Refillable = true;
        }

        #endregion

        private void Default_Setting()
        {
            IGR_APPROVED_CERTI.SetCellValue("PRINT_DATE", DateTime.Today );
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            int mIDX_SELECT_YN = IGR_APPROVED_CERTI.GetColumnToIndex("SELECT_YN");
            int mIDX_NAME = IGR_APPROVED_CERTI.GetColumnToIndex("NAME");
            int mIDX_CERT_TYPE = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_TYPE_NAME");
            int mIDX_PERPOSE = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_PRINT_PERPOSE");
            int mIDX_DESCRIPTION = IGR_APPROVED_CERTI.GetColumnToIndex("DESCRIPTION");

            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "A")
            {
                ibt_OK.Enabled = false;
                ibt_CANCEL.Enabled = false;

                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_PERPOSE].Updatable = 0;
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
            }
            else
            {
                if (mAPPROVE_STATE == "N")
                {
                    ibt_OK.Enabled = true;
                    ibt_CANCEL.Enabled = false;

                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_PERPOSE].Updatable = 1;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 1;
                }
                else
                {
                    ibt_OK.Enabled = false;
                    ibt_CANCEL.Enabled = true;

                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_NAME].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_CERT_TYPE].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_PERPOSE].Updatable = 0;
                    IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_DESCRIPTION].Updatable = 0;
                }
                IGR_APPROVED_CERTI.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
            SEARCH_DB();
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
            int vIDX_DUTY_PERIOD_ID = IGR_APPROVED_CERTI.GetColumnToIndex("CERT_PRINT_ID");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_APPROVED_CERTI.RowCount; i++)
            {
                if (iConv.ISNull(IGR_APPROVED_CERTI.GetCellValue(i, vIDX_SELECT_FLAG), "N") == "Y")
                {

                    idcAPPROVED.SetCommandParamValue("W_CERT_PRINT_ID", IGR_APPROVED_CERTI.GetCellValue(i, vIDX_DUTY_PERIOD_ID));
                    idcAPPROVED.SetCommandParamValue("W_APPROVE_STATUS", pApproved_Flag);
                    idcAPPROVED.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(idcAPPROVED.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(idcAPPROVED.GetCommandParamValue("O_MESSAGE"));
                    if (idcAPPROVED.ExcuteError || vSTATUS == "F")
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

        private void ILA_CERTIFICATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CERTIFICATE_W.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ILD_CERTIFICATE_W.SetLookupParamValue("W_WHERE", "HC.VALUE3 = 'Y'");
            ILD_CERTIFICATE_W.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_CERTIFICATE_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CERTIFICATE.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ILD_CERTIFICATE.SetLookupParamValue("W_WHERE", "HC.VALUE3 = 'Y'");
            ILD_CERTIFICATE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        private void ILA_SEARCH_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            ILD_SEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion



        private void ibt_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Update_Approve("A");
        }

        private void ibt_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Update_Approve("N");
        }

        private void APPROVED_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();
            
        }

        private void ILA_CERT_REMARK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CERT_REMARK.SetLookupParamValue("W_GROUP_CODE", "CERT_SEND");
            ILD_CERT_REMARK.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
    }
}