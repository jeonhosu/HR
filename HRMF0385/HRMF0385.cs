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

namespace HRMF0385
{
    public partial class HRMF0385 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0385(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            V_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            V_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void INIT_COLUMN()
        {
            IDA_MONTH_DAY_WEEK.Fill();
            if (IDA_MONTH_DAY_WEEK.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            int mGRID_START_COL = 8;   // 그리드 시작 COLUMN.
            int mIDX_Column;            // 시작 COLUMN.            
            int mMax_Column = 31;       // 종료 COLUMN.
            //int mENABLED_COLUMN;        // 사용여부 COLUMN.

            //object mENABLED_FLAG;       // 사용(표시)여부.
            object mCOLUMN_DESC;        // 헤더 프롬프트.

            for (mIDX_Column = 0; mIDX_Column < mMax_Column; mIDX_Column++)
            {
                //mENABLED_COLUMN = mMax_Column + mIDX_Column;
                //mENABLED_FLAG = IDA_MONTH_DAY_WEEK.CurrentRow[mENABLED_COLUMN];
                mCOLUMN_DESC = IDA_MONTH_DAY_WEEK.CurrentRow[mIDX_Column];
                if (iString.ISNull(mCOLUMN_DESC, "N") == "N".ToString())
                {
                    IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 0;
                }
                else
                {
                    IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                    IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                    IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);
                }
            }
            IGR_MONTH_DAILY_LIST.ResetDraw = true;
        }

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


        #endregion;


        #region ----- XL Print 1 Methods ----

        private void XLPrinting_Main()
        {            
            //print type 설정
            DialogResult vdlgResult;
            HRMF0385_PRINT_TYPE vHRMF0385_PRINT_TYPE = new HRMF0385_PRINT_TYPE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0385_PRINT_TYPE.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            string vPRINT_TYPE = iString.ISNull(vHRMF0385_PRINT_TYPE.Get_Printer_Type);
            if (vPRINT_TYPE == string.Empty)
            {
                return;
            }
            vHRMF0385_PRINT_TYPE.Dispose();

            string vSaveFileName = string.Empty;
            if (vPRINT_TYPE == "EXCEL")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "xls file(*.Xlsx)|*.Xlsx|(*.xls)|*.xls";
                vSaveFileDialog.DefaultExt = "Xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }
            else if (vPRINT_TYPE == "PDF")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "pdf file(*.pdf)|*.pdf";
                vSaveFileDialog.DefaultExt = "pdf";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }
             
            if (IGR_MONTH_DAILY_LIST.RowCount == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return;
            } 
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", iDate.ISMonth_Last(V_DUTY_YYYYMM.EditValue));
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0385");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            string vREPORT_FILE_NAME = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_FILE_NAME"));


            XLPrinting(vREPORT_FILE_NAME, vPRINT_TYPE, vSaveFileName); 
        }

        private void XLPrinting(string pReport_File_Name, string pOutput_Type, string pSaveFileName)
        {
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            string vMessageText = string.Empty;
            int vPageTotal = 0;
            int vPageNumber = 0;

            vMessageText = string.Format("Printing Starting", vPageTotal);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            //-------------------------------------------------------------------------------------
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface);
             
            try
            {
                //-------------------------------------------------------------------------------------
                if (pReport_File_Name == string.Empty)
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0385_001.xlsx";
                }
                else
                {
                    xlPrinting.OpenFileNameExcel = pReport_File_Name;
                }
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    object vObject;
                    int vCountRow = 0;

                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    xlPrinting.HeaderWrite(V_DUTY_YYYYMM.EditValue, vLOCAL_DATE, IGR_MONTH_DAILY_LIST);

                    vCountRow = IGR_MONTH_DAILY_LIST.RowCount;
                    if (vCountRow > 0)
                    {
                        vPageNumber = xlPrinting.LineWrite(IGR_MONTH_DAILY_LIST);
                    }

                    if (pOutput_Type == "PREVIEW")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PreView(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PRINT")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.Printing(1, vPageNumber);

                    }
                    else if (pOutput_Type == "PDF")
                    {//[PRINT]
                        ////xlPrinting.Printing(3, 4); //시작 페이지 번호, 종료 페이지 번호
                        xlPrinting.PDF(pSaveFileName);

                    }
                    else if (pOutput_Type == "EXCEL")
                    {
                        ////[SAVE]
                        xlPrinting.Save(pSaveFileName); //저장 파일명
                    }

                    vPageTotal = vPageTotal + vPageNumber;
                }
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();

                System.Windows.Forms.Application.UseWaitCursor = false;
                this.Cursor = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();

                return;
            }

            //-------------------------------------------------------------------------
            vMessageText = string.Format("Print End ^.^ [Tatal Page : {0}]", vPageTotal);
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }
         
        #endregion;


        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----      
        
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (V_CORP_ID.EditValue == null)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_CORP_ID.Focus();
                        return;
                    }

                    if (iString.ISNull(V_DUTY_YYYYMM.EditValue) == String.Empty)
                    {// 급여년월
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        V_DUTY_YYYYMM.Focus();
                        return;
                    }
                    INIT_COLUMN();
                    Application.DoEvents();

                    IDA_MONTH_DAILY_LIST.Fill();
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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_MONTH_DAILY_LIST);  
                }

            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0313_Load(object sender, EventArgs e)
        {
            IDA_MONTH_DAILY_LIST.FillSchema();

            // Year Month Setting
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", "2016-01");
            V_DUTY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            V_CORP_NAME.BringToFront();

            DefaultCorporation();
        }
       
        #endregion  

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_PERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", DateTime.Today);
        }

    }
}