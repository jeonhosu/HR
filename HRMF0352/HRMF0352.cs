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

namespace HRMF0352
{
    public partial class HRMF0352 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0352(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE_0.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            igbCORP_GROUP_0.Visible = false;

            if (iConv.ISNull(CORP_TYPE_0.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();

                irb_ALL_0.RadioButtonValue = "A";
                CORP_TYPE_0.EditValue = "A";
            }
            else if (iConv.ISNull(CORP_TYPE_0.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null && CORP_TYPE_0.EditValue.ToString() == "1")
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            //if (PERSON_ID_0.EditValue == null)
            //{// 성명
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=성명"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    NAME_0.Focus();
            //    return;
            //}

            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                IDA_DAY_COMMUTE.Fill();
                igrDAY_COMMUTE.Focus();
                SetRowColor(igrDAY_COMMUTE);
            }

            else if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                IDA_DAY_CLOSE_PERSON.Fill();
                igrDAY_LEAVE.Focus();
                SetRowColor(igrDAY_LEAVE);
            }
        }
        
        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {
            
            if (iConv.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }

            //idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pWork_Date);
            //idaDAY_HISTORY.Fill();
            //idaDUTY_PERIOD.Fill();
            //idaWORK_CALENDAR.Fill();
        }

        private void isSearch_Day_History(int pAdd_Day)
        {
           // idaDAY_HISTORY.Fill();
        }

        private void SetRowColor(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            int vIDX_SUMMARY_TYPE = pGrid.GetColumnToIndex("SUMMARY_TYPE");
            for(int r =0; r < pGrid.RowCount; r++)
            {
                if("T" == iConv.ISNull(pGrid.GetCellValue(r, vIDX_SUMMARY_TYPE)))
                {
                    pGrid.RowBackColor_Total(r);
                }
                else if ("P" == iConv.ISNull(pGrid.GetCellValue(r, vIDX_SUMMARY_TYPE)))
                {
                    pGrid.RowBackColor_Sum(r); 
                }
                else if ("W" == iConv.ISNull(pGrid.GetCellValue(r, vIDX_SUMMARY_TYPE)))
                {
                    pGrid.RowBackColor_SubSum(r);
                } 
            }
        }

        #endregion;

        #region ----- Excel Export II -----

        private void ExcelExport(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.RestoreDirectory = true;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Work Data";   //기본 파일명.수정필요.

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.Filter = "CSV File(*.csv)|*.csv|Excel file(*.xlsx)|*.xlsx|Excel file(*.xls)|*.xls";
            saveFileDialog1.DefaultExt = "xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
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
            int vCountRow = pAdapter.CurrentRows.Count;

            if (vCountRow < 1)
            {
                vMessageText = isMessageAdapter1.ReturnText("EAPP_10106");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);

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
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
                }
                else
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
                }

                //A new workbook is created.[Equivalent to creating a new workbook in MS Excel]
                //The new workbook will have 3 worksheets
                IWorkbook Exc_WorkBook = Exc_App.Workbooks.Create(1);
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel97to2003;
                }
                else
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel2007;
                }

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet sheet = Exc_WorkBook.Worksheets[0];

                //Export DataTable.
                sheet.ImportDataTable(pAdapter.OraDataTable(), false, 1, 1, pAdapter.CurrentRows.Count, pAdapter.OraSelectData.Columns.Count, true);

                //1.title insert
                int vHeaderCount = pGrid.GridAdvExColElement[0].HeaderElement.Count;
                for (int h = 1; h <= vHeaderCount; h++)
                {
                    sheet.InsertRow(h);
                    object vTitle = string.Empty;
                    for (int c = 0; c < pGrid.ColCount; c++)
                    {
                        if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL1_KR)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL1_KR;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL2_CN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL2_CN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL3_VN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL3_VN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL4_JP)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL4_JP;
                        }
                        else
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].Default;
                        }

                        sheet.Range[h, c + 1].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                        sheet.Range[h, c + 1].Value = iConv.ISNull(vTitle);
                        sheet.AutofitColumn(c + 1);
                        if (iConv.ISNull(pGrid.GridAdvExColElement[c].Visible) == "0")
                        {
                            sheet.HideColumn(c + 1);
                        }
                    }
                }

                ////2.prompt insert
                //sheet.InsertRow(2);
                //sheet.ImportDataTable(IDA_REJECT_DETAIL_TITLE.OraDataTable(), false, 2, 1); 
                //Exc_WorkBook.ActiveSheet.AutofitColumn(1);

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

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "CSV(*.csv)|*.csv";
            saveFileDialog.DefaultExt = ".csv";

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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
                
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT");  
                }

                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (TB_BASE.SelectedTab.TabIndex == 1)
                    {                        
                        ExcelExport(IDA_DAY_COMMUTE, igrDAY_COMMUTE);
                    }

                    else if (TB_BASE.SelectedTab.TabIndex == 2)
                    {
                        ExcelExport(IDA_DAY_CLOSE_PERSON, igrDAY_LEAVE); 
                    }
                    
                }

            }
        }
        #endregion;

        #region ----- XLPrinting -----

        private void XLPrinting(string pOutChoice)
        {
            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                //개인별출퇴근조회
                XLPrinting_1(pOutChoice);
            }
            else if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                //일근태마감조회
                XLPrinting_2(pOutChoice);
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XLPrinting1  -----

        private void XLPrinting_1(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object mCORP_NAME = string.Empty;
            object mPERIOD_DATE = string.Empty;
            object mPRINTED_DATE = string.Empty;
            //object mPRINTED_BY = string.Empty;

            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = igrDAY_COMMUTE.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("개인별출퇴근조회_{0}", DateTime.Today.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
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
            }
            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0352_001.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //if (iConv.ISNull(EMPLOYE_TYPE_0.EditValue) == "1")
                    //{
                    //    mTitle = GetPrompt(irbJOIN);
                    //}
                    //else if (iConv.ISNull(EMPLOYE_TYPE_0.EditValue) == "3")
                    //{
                    //    mTitle = GetPrompt(irbRETIRE);
                    //}
                    //else
                    //{
                    //    mTitle = "Non Title";
                    //}

                    // 헤더 인쇄.
                    IDC_PRINTED_VALUE.ExecuteNonQuery();
                    // 실제 인쇄
                    vPageNumber = xlPrinting.ExcelWrite(IDC_PRINTED_VALUE, igrDAY_COMMUTE, TB_BASE.SelectedTab.TabIndex);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- XLPrinting2  -----

        private void XLPrinting_2(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object mCORP_NAME = string.Empty;
            object mPERIOD_DATE = string.Empty;
            object mPRINTED_DATE = string.Empty;
            //object mPRINTED_BY = string.Empty;

            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = igrDAY_LEAVE.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("일근태마감조회_{0}", DateTime.Today.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
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
            }
            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vPageNumber = 0;
            //int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0352_002.xlsx";
                //-------------------------------------------------------------------------------------
                // 파일 오픈.
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    //if (iConv.ISNull(EMPLOYE_TYPE_0.EditValue) == "1")
                    //{
                    //    mTitle = GetPrompt(irbJOIN);
                    //}
                    //else if (iConv.ISNull(EMPLOYE_TYPE_0.EditValue) == "3")
                    //{
                    //    mTitle = GetPrompt(irbRETIRE);
                    //}
                    //else
                    //{
                    //    mTitle = "Non Title";
                    //}

                    // 헤더 인쇄.
                    IDC_PRINTED_VALUE.ExecuteNonQuery();
                    // 실제 인쇄
                    vPageNumber = xlPrinting.ExcelWrite(IDC_PRINTED_VALUE, igrDAY_LEAVE, TB_BASE.SelectedTab.TabIndex);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }
        #endregion;

      

        #region ----- Form Event -----

        private void HRMF0352_Load(object sender, EventArgs e)
        {
            this.Visible = true;

            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = DateTime.Today;
            
            // CORP SETTING
            DefaultCorporation();

            // idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            // idcDEFAULT_VALUE.ExecuteNonQuery();
            IDA_DAY_COMMUTE.FillSchema();
            IDA_DAY_CLOSE_PERSON.FillSchema();  
        }

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_0.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_LEAVE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_LEAVE.GetCellValue("PERSON_ID"), igrDAY_LEAVE.GetCellValue("WORK_DATE"));
        }

        private void idaDAY_LEAVE_PERSON_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_LEAVE.GetCellValue("PERSON_ID"), igrDAY_LEAVE.GetCellValue("WORK_DATE"));
        }

        #endregion

        #region ----- LookUp Event -----

        private void ildHOLY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaJOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", END_DATE_0.EditValue);
        }

        private void ilaFLOOR_0_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        #endregion

    }
}