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

namespace SOMF0403
{
    public partial class SOMF0403_EXPORT : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime(); 

        #endregion;

        #region ----- Constructor -----

        public SOMF0403_EXPORT()
        {
            InitializeComponent();
        }

        public SOMF0403_EXPORT(Form pMainForm, ISAppInterface pAppInterface
                                , object pFCST_HEADER_ID, object pFCST_WEEK_NO, object pFCST_WEEK_DATE
                                , object pCUSTOMER_ID, object pCUSTOMER_CODE, object pCUSTOMER_DESC 
                                , object pINVENTORY_ITEM_ID, object pITEM_CODE, object pITEM_DESC
                                , object pITEM_SECTION_CODE, object pITEM_SECTION_DESC
                                , object pEXCHANGE_RATE_TYPE
                                , object pORDER_FLAG, object pREGISTER_FLAG, object pDISCONTINUED_FLAG
                                , object pSALES_PERSON_ID, object pSALES_PERSON_NUM, object pSALES_PERSON_NAME)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            FCST_HEADER_ID.EditValue = pFCST_HEADER_ID;
            FCST_WEEK_NO.EditValue = pFCST_WEEK_NO;
            FCST_WEEK_DATE.EditValue = pFCST_WEEK_DATE;

            V_CUSTOMER_ID.EditValue = pCUSTOMER_ID;
            V_CUSTOMER_CODE.EditValue = pCUSTOMER_CODE;
            V_CUSTOMER_DESC.EditValue = pCUSTOMER_DESC;
            V_ITEM_ID.EditValue = pINVENTORY_ITEM_ID;
            V_ITEM_CODE.EditValue = pITEM_CODE;
            V_ITEM_DESC.EditValue = pITEM_DESC;
            V_ITEM_SECTION_CODE.EditValue = pITEM_SECTION_CODE;
            V_ITEM_SECTION_DESC.EditValue = pITEM_SECTION_DESC;
            V_EXCHANGE_RATE_TYPE.EditValue = pEXCHANGE_RATE_TYPE;
            V_ORDER_FLAG.CheckBoxValue = pORDER_FLAG;
            V_REGISTER_FLAG.CheckBoxValue = pREGISTER_FLAG;
            V_DISCONTINUED_FLAG.CheckBoxValue = pDISCONTINUED_FLAG;
            V_SALES_PERSON_ID.EditValue = pSALES_PERSON_ID;
            V_SALES_PERSON_NUM.EditValue = pSALES_PERSON_NUM;
            V_SALES_PERSON_NAME.EditValue = pSALES_PERSON_NAME;
        }

        #endregion;

        #region ----- Private Methods ----

        private bool Search_DB()
        {
            PT_MESSAGE.PromptTextElement[0].Default = "Data Searching.. Waiting please";
            Application.DoEvents();
            try
            {                  
                IDA_EXPORT_SALES_FCST_LINE.Fill();
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
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


        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx vGrid)
        {
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.Filter = "Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xls";

            PT_MESSAGE.PromptTextElement[0].Default = "Data Exporting.. Waiting please";
            Application.DoEvents();
 
            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();
 
                vExport.GridToExcel(vGrid.BaseGrid.Model, vSaveFileDialog.FileName,
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

            this.DialogResult = DialogResult.OK;
            this.Close();
            return; 
        }
          
        private void Xls_Export()
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Sales FCST List";

            SaveFileDialog vSaveFileDialog = new SaveFileDialog();
            vSaveFileDialog.Title = "Excel Save";
            vSaveFileDialog.RestoreDirectory = true;
            vSaveFileDialog.FileName = vSaveFileName;
            vSaveFileDialog.Filter = "Excel file(*.xlsx)|*.xlsx|Excel file(*.xls)|*.xls";
            vSaveFileDialog.DefaultExt = "xlsx";
            if (vSaveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = vSaveFileDialog.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                try
                {
                    if (vFileName.Exists)
                    {
                        vFileName.Delete();
                    }
                }
                catch (Exception EX)
                {
                    MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            vMessageText = string.Format(" Writing Starting...");

            System.Windows.Forms.Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //DATA 조회    
            IDA_EXPORT_SALES_FCST_LINE.Fill();
            int vCountRow = IDA_EXPORT_SALES_FCST_LINE.CurrentRows.Count;

            if (vCountRow < 1)
            {
                PT_MESSAGE.PromptTextElement[0].Default = isMessageAdapter1.ReturnText("EAPP_10106");
                PT_MESSAGE.Refresh();
                 
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            try
            {
                //Step 1 : Instantiate the spreadsheet creation engine.
                ExcelEngine ExcelEngine = new ExcelEngine();

                //Step 2 : Instantiate the excel application object.
                IApplication Exc_App = ExcelEngine.Excel;

                //set 2.1 : file Extension check =>xlsx, xls 
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLSX")
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
                }
                else
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
                }

                //A new workbook is created.[Equivalent to creating a new workbook in MS Excel]
                //The new workbook will have 3 worksheets
                IWorkbook Exc_WorkBook = Exc_App.Workbooks.Create(1);
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLSX")
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel2007;
                }
                else
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel97to2003;
                }

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet sheet = Exc_WorkBook.Worksheets[0];

                //Export DataTable.
                sheet.ImportDataTable(IDA_EXPORT_SALES_FCST_LINE.OraDataTable(), false, 1, 1, IDA_EXPORT_SALES_FCST_LINE.CurrentRows.Count, IDA_EXPORT_SALES_FCST_LINE.OraSelectData.Columns.Count);

                //1.title insert
                int vHeaderRow = IGR_SALES_FCST_LINE.ColHeaderCount;
                for (int h = 1; h <= vHeaderRow; h++)
                {
                    sheet.InsertRow(h);
                    for (int c = 0; c < IGR_SALES_FCST_LINE.ColCount; c++)
                    {
                        sheet.Range[h, c + 1].Value = IGR_SALES_FCST_LINE.GridAdvExColElement[c].HeaderElement[(vHeaderRow - h)].Default;
                    }
                }

                for (int c = 0; c < IGR_SALES_FCST_LINE.ColCount; c++)
                {
                    sheet.AutofitColumn(c + 1);
                }

                //IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[1].TL1_KR = (mCOLUMN_DESC);
                //IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[1].Default = (mCOLUMN_DESC);

                //sheet.Range[1, 1].Value = vTRX_PERIOD_DATE;
                //2.prompt insert
                //sheet.InsertRow(2);
                //sheet.ImportDataTable( (IDA_REJECT_DETAIL_TITLE.OraDataTable(), false, 2, 1);

                //Saving the workbook to disk.
                Exc_WorkBook.SaveAs(vSaveFileName);

                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(vSaveFileName);
                }

            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }
 
        #endregion

        #region ----- Excel Upload -----

        //private bool Excel_Upload()
        //{
        //    if (iConv.ISNull(FCST_HEADER_ID.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        FCST_WEEK_NO.Focus();
        //        return false;
        //    }

        //    DateTime vFCST_WEEK_DATE = iDate.ISGetDate(FCST_WEEK_DATE.EditValue);
        //    string vSTATUS = string.Empty;
        //    string vMESSAGE = string.Empty;
        //    bool vXL_Load_OK = false;
            
        //    if (iConv.ISNull(FILE_PATH.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return false;
        //    }
        //    Application.UseWaitCursor = true;
        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    string vOPenFileName = FILE_PATH.EditValue.ToString();
        //    XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

        //    try
        //    {
        //        vXL_Upload.OpenFileName = vOPenFileName;
        //        vXL_Load_OK = vXL_Upload.OpenXL();
        //    }
        //    catch (Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return false;
        //    }

        //    // 업로드 아답터 fill //
        //    IDA_EXPORT_SALES_FCST_LINE.Cancel();
        //    IDA_EXPORT_SALES_FCST_LINE.Fill();
        //    try
        //    {
        //        if (vXL_Load_OK == true)
        //        {
        //            vXL_Load_OK = vXL_Upload.LoadXL(IDA_EXPORT_SALES_FCST_LINE, vFCST_WEEK_DATE, 3);
        //            if (vXL_Load_OK == false)
        //            {
        //                IDA_EXPORT_SALES_FCST_LINE.Cancel();
        //            }
        //            else
        //            {
        //                IDA_EXPORT_SALES_FCST_LINE.Update();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        IDA_EXPORT_SALES_FCST_LINE.Cancel();
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //        vXL_Upload.DisposeXL();

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return false;
        //    }
        //    vXL_Upload.DisposeXL();
            
        //    Application.UseWaitCursor = false;
        //    this.Cursor = Cursors.Default;
        //    Application.DoEvents();

        //    return true;            
        //}

        //private bool Excel_Upload_Loop()
        //{
        //    if (iConv.ISNull(FCST_HEADER_ID.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        FCST_WEEK_NO.Focus();
        //        return false;
        //    }

        //    DateTime vFCST_WEEK_DATE = iDate.ISGetDate(FCST_WEEK_DATE.EditValue);
        //    string vSTATUS = string.Empty;
        //    string vMESSAGE = string.Empty;

        //    System.Type vType = null;

        //    if (iConv.ISNull(FILE_PATH.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return false;
        //    }
        //    Application.UseWaitCursor = true;
        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    string vOPenFileName = FILE_PATH.EditValue.ToString();

        //    //--------------------------------------------------------------------------------------
        //    //excel 개체 생성
        //    //Step 1 : Instantiate the spreadsheet creation engine.
        //    ExcelEngine ExcelEngine = new ExcelEngine();
        //    //Step 2 : Instantiate the excel application object.
        //    IApplication Exc_App = ExcelEngine.Excel;

        //    if (Path.GetExtension(vOPenFileName).ToUpper() == ".XLSX")
        //    {
        //        Exc_App.DefaultVersion = ExcelVersion.Excel2007;
        //    }
        //    else
        //    {
        //        Exc_App.DefaultVersion = ExcelVersion.Excel97to2003;
        //    }

        //    //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
        //    //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
        //    //IWorkbook workbook = application.Workbooks.Open(@"..\..\..\..\..\..\..\..\..\Common\Data\XlsIO\NorthwindDataTemplate.xls");
        //    IWorkbook Exc_WorkBook = null;

        //    try
        //    {
        //        //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
        //        //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
        //        //IWorkbook workbook = application.Workbooks.Open(@"..\..\..\..\..\..\..\..\..\Common\Data\XlsIO\NorthwindDataTemplate.xls");
        //        Exc_WorkBook = Exc_App.Workbooks.Open(@vOPenFileName, ExcelOpenType.Automatic);
        //        //Exc_WorkBook = ExcelEngine.Excel.Workbooks.Open(@vOPenFileName, ExcelOpenType.Automatic);
        //    }
        //    catch (Exception Ex)
        //    {
        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();

        //        //No exception will be thrown if there are unsaved workbooks.
        //        ExcelEngine.ThrowNotSavedOnDestroy = false;
        //        ExcelEngine.Dispose();

        //        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        return false;
        //    }
        //    try
        //    {
        //        //The first worksheet object in the worksheets collection is accessed.
        //        IWorksheet Exc_Sheet = Exc_WorkBook.Worksheets[0];

        //        //Read data from spreadsheet.
        //        DataTable customersTable = Exc_Sheet.ExportDataTable(Exc_Sheet.Range[2,1, Exc_Sheet.Range.End.Row, Exc_Sheet.Range.End.Column], ExcelExportDataTableOptions.ColumnNames);

        //        int vTotalRow = customersTable.Rows.Count;
        //        int vRowCount = 0;

        //        foreach (System.Data.DataRow vRow in customersTable.Rows)
        //        {
        //            vRowCount++;
        //            if (iConv.ISNull(vRow[0]) != String.Empty)
        //            {
        //                IDA_EXPORT_SALES_FCST_LINE.AddUnder();

        //                for (int vCol = 0; vCol < customersTable.Columns.Count; vCol++)
        //                {
        //                    //vType = IDA_UPLOAD_FCST_LINE.CurrentRow.Table.Columns[vCol].DataType;
        //                    //if (vType.Name == "Decimal" || vType.Name == "Double")
        //                    //{
        //                    //    IDA_UPLOAD_FCST_LINE.CurrentRow[vCol] = iConv.ISDecimaltoZero(vRow[vCol], 0);
        //                    //}
        //                    //else
        //                    //{
        //                    //    IDA_UPLOAD_FCST_LINE.CurrentRow[vCol] = vRow[vCol];
        //                    //}
        //                    IDA_EXPORT_SALES_FCST_LINE.CurrentRow[vCol] = vRow[vCol];
        //                }
        //            }
        //            PB_UPLOAD.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vTotalRow)) * 100F;
        //            Application.DoEvents(); 
        //        }
        //    }
        //    catch (Exception Ex)
        //    {
        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();

        //        //Close the workbook.
        //        Exc_WorkBook.Close();

        //        //No exception will be thrown if there are unsaved workbooks.
        //        ExcelEngine.ThrowNotSavedOnDestroy = false;
        //        ExcelEngine.Dispose();

        //        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        return false;
        //    }
        //    //Close the workbook.
        //    Exc_WorkBook.Close();

        //    //No exception will be thrown if there are unsaved workbooks.
        //    ExcelEngine.ThrowNotSavedOnDestroy = false;
        //    ExcelEngine.Dispose();

        //    //XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

        //    //try
        //    //{
        //    //    vXL_Upload.OpenFileName = vOPenFileName;
        //    //    vXL_Load_OK = vXL_Upload.OpenXL();
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //    //    Application.UseWaitCursor = false;
        //    //    this.Cursor = Cursors.Default;
        //    //    Application.DoEvents();
        //    //    return false;
        //    //}

        //    //// 업로드 아답터 fill //
        //    //IDA_UPLOAD_FCST_LINE.Cancel();
        //    //IDA_UPLOAD_FCST_LINE.Fill();
        //    //try
        //    //{
        //    //    if (vXL_Load_OK == true)
        //    //    {
        //    //        vXL_Load_OK = vXL_Upload.LoadXL(IDA_UPLOAD_FCST_LINE, vFCST_WEEK_DATE, 3);
        //    //        if (vXL_Load_OK == false)
        //    //        {
        //    //            IDA_UPLOAD_FCST_LINE.Cancel();
        //    //        }
        //    //        else
        //    //        {
        //    //            IDA_UPLOAD_FCST_LINE.Update();
        //    //        }
        //    //    }
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    IDA_UPLOAD_FCST_LINE.Cancel();
        //    //    isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //    //    vXL_Upload.DisposeXL();

        //    //    Application.UseWaitCursor = false;
        //    //    this.Cursor = Cursors.Default;
        //    //    Application.DoEvents();
        //    //    return false;
        //    //}
        //    //vXL_Upload.DisposeXL();

        //    Application.UseWaitCursor = false;
        //    this.Cursor = Cursors.Default;
        //    Application.DoEvents();

        //    return true;            
        //}

        #endregion

        private void INIT_FCST_COLUMN(object pFCST_WEEK_DATE)
        {
            IGR_SALES_FCST_LINE.Focus();

            int mGRID_START_COL = 23;    // 그리드 시작 COLUMN INDEX.
            int mIDX_Column = 0;        // 시작 COLUMN.         

            string mCOLUMN_DESC;        // 헤더 프롬프트.
            string mWEEK;               // 요일코드

            //헤더 0.
            IDA_PROMPT_GRID.SetSelectParamValue("W_FCST_WEEK_DATE", pFCST_WEEK_DATE); 
            IDA_PROMPT_GRID.Fill();
            if (IDA_PROMPT_GRID.OraSelectData.Rows.Count == 0)
            {
                return;
            }

            foreach (DataRow vRow in IDA_PROMPT_GRID.OraSelectData.Rows)
            {
                mCOLUMN_DESC = iConv.ISNull(vRow["FCST_ITEM_DATE"]);
                mWEEK = iConv.ISNull(vRow["FCST_ITEM_WEEK"]);

                if (mCOLUMN_DESC == string.Empty)
                {
                    IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column + 1].Visible = 0; 
                }
                else
                {
                    IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column + 1].Visible = 1;

                    IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].TL1_KR = (mCOLUMN_DESC);
                    IGR_SALES_FCST_LINE.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].Default = (mCOLUMN_DESC);
                }
                mIDX_Column++;
            }

            IGR_SALES_FCST_LINE.LastConfirmChanges();
            IGR_SALES_FCST_LINE.ResetDraw = true;
        }


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

        private void SOMF0403_EXPORT_Load(object sender, EventArgs e)
        {
            PT_MESSAGE.PromptTextElement[0].Default = ""; 
        }

        private void SOMF0403_EXPORT_Shown(object sender, EventArgs e)
        {
            INIT_FCST_COLUMN(FCST_WEEK_DATE.EditValue);              
        }

        private void BTN_UPLOAD_EXCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (Search_DB() == false)
            //{            
            //    return;
            //}             
            Xls_Export();
            //ExcelExport(IGR_SALES_FCST_LINE); 
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
         
        #endregion

    }
}