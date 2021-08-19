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

namespace HRMF0634
{
    public partial class HRMF0634 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0634()
        {
            InitializeComponent();
        }

        public HRMF0634(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        public HRMF0634(Form pMainForm, ISAppInterface pAppInterface, object pJOB_NO)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_DEPT_NAME.EditValue = pJOB_NO;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            IGR_RETIRE_RESERVE_DC.LastConfirmChanges();
            IDA_RETIRE_RESERVE_DC.OraSelectData.AcceptChanges();
            IDA_RETIRE_RESERVE_DC.Refillable = true;

            IDA_RETIRE_RESERVE_DC.Fill();
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
                    IDA_RETIRE_RESERVE_DC.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_RETIRE_RESERVE_DC.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_RETIRE_RESERVE_DC);
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


        private void W_JOB_NO_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void HRMF0634_Load(object sender, EventArgs e)
        {            
            W_START_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_END_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            DefaultCorporation();              //Default Corp.

            IDA_RETIRE_RESERVE_DC.FillSchema();
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void IGR_OPERATION_CAPA_Click(object sender, EventArgs e)
        {

        }

        private void IGR_OPERATION_CAPA_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
        

        }

        #region ----- Form Event ------



        #endregion

        #region ----- Lookup Event ------

        #endregion

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ibt_Calculation_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_START_YYYYMM.EditValue) == string.Empty)
            {
                return;
            }

            DialogResult dlgResult;
            HRMF0634_CAL vHRMF0634_CAL = new HRMF0634_CAL( isAppInterfaceAdv1.AppInterface
                                                        , W_START_YYYYMM.EditValue
                                                        , "CAL"
                                                        , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                        , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                        , W_FLOOR_ID.EditValue, W_FLOOR_NAME.EditValue
                                                        , icb_YEAR_UNDER.CheckBoxValue
                                                        , W_PERSON_ID.EditValue, W_NAME.EditValue
                                                        );

            dlgResult = vHRMF0634_CAL.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
            }
            vHRMF0634_CAL.Dispose();
            IDA_RETIRE_RESERVE_DC.Fill();
        }

        private void isButton2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (iString.ISNull(W_START_YYYYMM.EditValue) == string.Empty)
            //{
            //    return;
            //}
            DialogResult dlgResult;
            HRMF0634_CAL vHRMF0634_CAL = new HRMF0634_CAL(isAppInterfaceAdv1.AppInterface
                                                        , W_START_YYYYMM.EditValue
                                                        , "CLOSE"
                                                        , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                        , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                        , W_FLOOR_ID.EditValue, W_FLOOR_NAME.EditValue
                                                        , icb_YEAR_UNDER.CheckBoxValue
                                                        , W_PERSON_ID.EditValue, W_NAME.EditValue
                                                        );


            dlgResult = vHRMF0634_CAL.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
            }
            vHRMF0634_CAL.Dispose();

            IDA_RETIRE_RESERVE_DC.Fill();

        }

        private void ILA_START_YYYYMM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 3)));
        }

        private void ILA_END_YYYYMM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 3)));
        }
    }
}