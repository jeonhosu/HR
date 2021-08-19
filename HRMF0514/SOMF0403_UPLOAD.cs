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
using Syncfusion.XlsIO;

namespace SOMF0403
{
    public partial class SOMF0403_UPLOAD : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public float Set_BarFillPercent
        {
            get
            {
                return PB_UPLOAD.BarFillPercent;
            }
            set
            {
                PB_UPLOAD.BarFillPercent = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public SOMF0403_UPLOAD()
        {
            InitializeComponent();
        }

        public SOMF0403_UPLOAD(Form pMainForm, ISAppInterface pAppInterface
                                , object pFCST_HEADER_ID, object pFCST_WEEK_NO, object pFCST_WEEK_DATE)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            FCST_HEADER_ID.EditValue = pFCST_HEADER_ID;
            FCST_WEEK_NO.EditValue = pFCST_WEEK_NO;
            FCST_WEEK_DATE.EditValue = pFCST_WEEK_DATE;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {

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

        #endregion;

        #region ----- Excel Upload -----

        private bool Excel_Upload()
        {
            if (iConv.ISNull(FCST_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                FCST_WEEK_NO.Focus();
                return false;
            }

            DateTime vFCST_WEEK_DATE = iDate.ISGetDate(FCST_WEEK_DATE.EditValue);
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            bool vXL_Load_OK = false;
            
            if (iConv.ISNull(FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vOPenFileName = FILE_PATH.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return false;
            }

            // 업로드 아답터 fill //
            IDA_UPLOAD_FCST_LINE.Cancel();
            IDA_UPLOAD_FCST_LINE.Fill();
            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDA_UPLOAD_FCST_LINE, vFCST_WEEK_DATE, 3);
                    if (vXL_Load_OK == false)
                    {
                        IDA_UPLOAD_FCST_LINE.Cancel();
                    }
                    else
                    {
                        IDA_UPLOAD_FCST_LINE.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_UPLOAD_FCST_LINE.Cancel();
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return false;
            }
            vXL_Upload.DisposeXL();
            
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            return true;            
        }


        private bool Excel_Upload_Loop()
        {
            if (iConv.ISNull(FCST_HEADER_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                FCST_WEEK_NO.Focus();
                return false;
            }

            DateTime vFCST_WEEK_DATE = iDate.ISGetDate(FCST_WEEK_DATE.EditValue);
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;

            System.Type vType = null;

            if (iConv.ISNull(FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vOPenFileName = FILE_PATH.EditValue.ToString();

            //--------------------------------------------------------------------------------------
            //excel 개체 생성
            //Step 1 : Instantiate the spreadsheet creation engine.
            ExcelEngine ExcelEngine = new ExcelEngine();
            //Step 2 : Instantiate the excel application object.
            IApplication Exc_App = ExcelEngine.Excel;

            if (Path.GetExtension(vOPenFileName).ToUpper() == ".XLSX")
            {
                Exc_App.DefaultVersion = ExcelVersion.Excel2007;
            }
            else
            {
                Exc_App.DefaultVersion = ExcelVersion.Excel97to2003;
            }

            //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
            //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
            //IWorkbook workbook = application.Workbooks.Open(@"..\..\..\..\..\..\..\..\..\Common\Data\XlsIO\NorthwindDataTemplate.xls");
            IWorkbook Exc_WorkBook = null;

            try
            {
                //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
                //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
                //IWorkbook workbook = application.Workbooks.Open(@"..\..\..\..\..\..\..\..\..\Common\Data\XlsIO\NorthwindDataTemplate.xls");
                Exc_WorkBook = Exc_App.Workbooks.Open(@vOPenFileName, ExcelOpenType.Automatic);
                //Exc_WorkBook = ExcelEngine.Excel.Workbooks.Open(@vOPenFileName, ExcelOpenType.Automatic);
            }
            catch (Exception Ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }
            try
            {
                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet Exc_Sheet = Exc_WorkBook.Worksheets[0];

                //Read data from spreadsheet.
                DataTable customersTable = Exc_Sheet.ExportDataTable(Exc_Sheet.Range[2,1, Exc_Sheet.Range.End.Row, Exc_Sheet.Range.End.Column], ExcelExportDataTableOptions.ColumnNames);

                int vTotalRow = customersTable.Rows.Count;
                int vRowCount = 0;

                foreach (System.Data.DataRow vRow in customersTable.Rows)
                {
                    vRowCount++;
                    if (iConv.ISNull(vRow[0]) != String.Empty)
                    {
                        IDA_UPLOAD_FCST_LINE.AddUnder();

                        for (int vCol = 0; vCol < customersTable.Columns.Count; vCol++)
                        {
                            //vType = IDA_UPLOAD_FCST_LINE.CurrentRow.Table.Columns[vCol].DataType;
                            //if (vType.Name == "Decimal" || vType.Name == "Double")
                            //{
                            //    IDA_UPLOAD_FCST_LINE.CurrentRow[vCol] = iConv.ISDecimaltoZero(vRow[vCol], 0);
                            //}
                            //else
                            //{
                            //    IDA_UPLOAD_FCST_LINE.CurrentRow[vCol] = vRow[vCol];
                            //}
                            IDA_UPLOAD_FCST_LINE.CurrentRow[vCol] = vRow[vCol];
                        }
                    }
                    PB_UPLOAD.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vTotalRow)) * 100F;
                    Application.DoEvents(); 
                }
            }
            catch (Exception Ex)
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }
            //Close the workbook.
            Exc_WorkBook.Close();

            //No exception will be thrown if there are unsaved workbooks.
            ExcelEngine.ThrowNotSavedOnDestroy = false;
            ExcelEngine.Dispose();

            //XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

            //try
            //{
            //    vXL_Upload.OpenFileName = vOPenFileName;
            //    vXL_Load_OK = vXL_Upload.OpenXL();
            //}
            //catch (Exception ex)
            //{
            //    isAppInterfaceAdv1.OnAppMessage(ex.Message);

            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();
            //    return false;
            //}

            //// 업로드 아답터 fill //
            //IDA_UPLOAD_FCST_LINE.Cancel();
            //IDA_UPLOAD_FCST_LINE.Fill();
            //try
            //{
            //    if (vXL_Load_OK == true)
            //    {
            //        vXL_Load_OK = vXL_Upload.LoadXL(IDA_UPLOAD_FCST_LINE, vFCST_WEEK_DATE, 3);
            //        if (vXL_Load_OK == false)
            //        {
            //            IDA_UPLOAD_FCST_LINE.Cancel();
            //        }
            //        else
            //        {
            //            IDA_UPLOAD_FCST_LINE.Update();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    IDA_UPLOAD_FCST_LINE.Cancel();
            //    isAppInterfaceAdv1.OnAppMessage(ex.Message);
            //    vXL_Upload.DisposeXL();

            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();
            //    return false;
            //}
            //vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            return true;            
        }


        #endregion

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                   
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
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void SOMF0403_UPLOAD_Load(object sender, EventArgs e)
        {

            IDA_UPLOAD_FCST_LINE.FillSchema();
        }

        private void SOMF0403_UPLOAD_Shown(object sender, EventArgs e)
        {
            PB_UPLOAD.BarFillPercent = 0;
        }

        private void BTN_FILE_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Upload File";
                openFileDialog1.Filter = "Excel File(*.xls)|*.xls";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private void BTN_UPLOAD_EXCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (Excel_Upload() == false)
            //{
            //    this.DialogResult = DialogResult.No;
            //    return;
            //}

            if (Excel_Upload_Loop() == false)
            {
                this.DialogResult = DialogResult.No;
                return;
            }
            else
            {
                PT_MESSAGE.PromptTextElement[0].Default = "Excel Import Completed... Wait Please";
                Application.DoEvents();

                PT_MESSAGE.PromptTextElement[0].Default = "DB Saving... Wait Please";
                Application.DoEvents();
                try
                {
                    IDA_UPLOAD_FCST_LINE.Update();
                }
                catch  
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            }
            this.DialogResult = DialogResult.Yes;
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        
        #endregion

        #region ----- Adapter Event -----


        private void IDA_UPLOAD_FCST_LINE_UpdateCompleted(object pSender)
        {
            IDC_SET_DBMS_MVIEW_REFRESH.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_DBMS_MVIEW_REFRESH.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_DBMS_MVIEW_REFRESH.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_DBMS_MVIEW_REFRESH.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }            
        }

        #endregion

    }
}