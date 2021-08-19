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

namespace HRMF0202
{
    public partial class HRMF0202 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime ISDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0202(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
           // isAppInterfaceAdv1.AppInterface.Attribute_A = "4";
            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE_1.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
                CORP_TYPE_2.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
                CORP_TYPE_3.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }

        }

        #endregion;

        #region ----- Property Method ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP_0.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            idcDEFAULT_CORP_0.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP_0.ExecuteNonQuery();

            CORP_NAME_1.BringToFront();
            CORP_NAME_2.BringToFront();
            CORP_NAME_3.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            igr_CHECK_BOX2.BringToFront();
            igr_CHECK_BOX3.BringToFront();

            if (iConv.ISNull(CORP_TYPE_1.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();
                igr_CHECK_BOX2.Visible = true; //.Show();
                igr_CHECK_BOX3.Visible = true; //.Show();

                irb_ALL_0.RadioButtonValue = "A";
                irb_ALL_1.RadioButtonValue = "A";
                irb_ALL_2.RadioButtonValue = "A";
                CORP_TYPE_1.EditValue = "A";
                CORP_TYPE_2.EditValue = "A";
                CORP_TYPE_3.EditValue = "A";
            }
            else if(iConv.ISNull(CORP_TYPE_1.EditValue) == "1")
            {
                CORP_NAME_1.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_1.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_2.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_2.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_3.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_3.EditValue = idcDEFAULT_CORP_0.GetCommandParamValue("O_CORP_ID");
            }

            
        }

        private void DefaultEmploye()
        {
            idcDEFAULT_EMPLOYE_TYPE_3.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            idcDEFAULT_EMPLOYE_TYPE_3.ExecuteNonQuery();
            EMPLOYE_TYPE_NAME_3.EditValue = idcDEFAULT_EMPLOYE_TYPE_3.GetCommandParamValue("O_CODE_NAME");
            EMPLOYE_TYPE_3.EditValue = idcDEFAULT_EMPLOYE_TYPE_3.GetCommandParamValue("O_CODE");
        }

        private void DefaultDateTime()
        {
            STD_DATE_1.EditValue = DateTime.Today;

            DATE_FR_2.EditValue = ISDate.ISMonth_1st(System.DateTime.Today);
            DATE_TO_2.EditValue = ISDate.ISMonth_Last(System.DateTime.Today);
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);            
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SEARCH_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                idaPERSON_DETAIL_DAY.Fill();
                IGR_PERSON_DATE.Focus();

            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                idaPERSON_DETAIL_PERIOD.Fill();
                IGR_PERSON_PERIOD.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 3)
            {
                idaPERSON_DETAIL.Fill();
                IGR_PERSON_DETAIL.Focus();
            }
        }

        #endregion

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

                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);

               

                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

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


        #region ----- Excel Export II -----

        private void Xls_Export(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.RestoreDirectory = true;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Person List";

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

                        sheet.Range[1, c + 1].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                        sheet.Range[1, c + 1].Value = iConv.ISNull(vTitle);
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

        #region ----- MDi ToolBar Button Evetn -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
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
                   
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPERSON_DETAIL.IsFocused)
                    {
                        idaPERSON_DETAIL.Cancel();
                    }
                    else if (idaPERSON_DETAIL_DAY.IsFocused)
                    {
                        idaPERSON_DETAIL_DAY.Cancel();
                    }
                    else if (idaPERSON_DETAIL_PERIOD.IsFocused)
                    {
                        idaPERSON_DETAIL_PERIOD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (TB_MAIN.SelectedTab.TabIndex == TP_PERSON_DATE.TabIndex)
                    {
                        Xls_Export(idaPERSON_DETAIL_DAY, IGR_PERSON_DATE); 
                    }
                    else if(TB_MAIN.SelectedTab.TabIndex == TP_PERSON_PERIOD.TabIndex)
                    {
                        Xls_Export(idaPERSON_DETAIL_PERIOD, IGR_PERSON_PERIOD);
                    }
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_PERSON_DETAIL.TabIndex)
                    {
                        Xls_Export(idaPERSON_DETAIL, IGR_PERSON_DETAIL);
                    }
                }
            }
        }

        #endregion

        #region ----- Form Event -----

        private void HRMF0202_Load(object sender, EventArgs e)
        {
                      
        }

        private void HRMF0202_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();
            DefaultEmploye();
            DefaultDateTime();

            RB_PERSON_NUM_1.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SORT_TYPE_1.EditValue = RB_PERSON_NUM_1.RadioCheckedString;

            RB_PERSON_NUM_2.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SORT_TYPE_2.EditValue = RB_PERSON_NUM_2.RadioCheckedString;

            RB_PERSON_NUM_3.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SORT_TYPE_3.EditValue = RB_PERSON_NUM_3.RadioCheckedString;
        }

        private void isTAB_Click(object sender, EventArgs e)
        {
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                RB_PERSON_NUM_1.CheckedState = ISUtil.Enum.CheckedState.Checked;
                W_SORT_TYPE_1.EditValue = RB_PERSON_NUM_1.RadioCheckedString;

            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {

            }
            else if (TB_MAIN.SelectedTab.TabIndex == 3)
            {

            }
            SEARCH_DB();
        }

        private void RB_PERSON_NUM_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;

            W_SORT_TYPE_1.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void RB_PERSON_NUM_2_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;

            W_SORT_TYPE_2.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void RB_PERSON_NUM_3_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;

            W_SORT_TYPE_3.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ilaOPERATING_UNIT_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", CORP_ID_1.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ilaOPERATING_UNIT_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", CORP_ID_2.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "N");
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_TYPE", CORP_TYPE_2.EditValue);
        }

        private void ilaOPERATING_UNIT_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", CORP_ID_3.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "N");
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_TYPE", CORP_TYPE_3.EditValue);
        }

        private void ilaOPERATING_UNIT_4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", CORP_ID_1.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "N");
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_TYPE", CORP_TYPE_1.EditValue);
        }

        private void ILA_JOB_CATEGORY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        private void ILA_CONTRACT_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_TYPE", "Y");
        }

        private void ilaDEPT_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaFLOOR_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_JOB_CATEGORY_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        private void ILA_CONTRACT_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_TYPE", "Y");
        }

        private void ilaDEPT_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaFLOOR_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_JOB_CATEGORY_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        private void ilaEMPLOYE_TYPE_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EMPLOYE_TYPE", "Y");
        }

        private void ILA_CONTRACT_TYPE_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("CONTRACT_TYPE", "Y");
        }

        private void irb_CORP_TYPE1_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_1.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void isRadioButtonAdv3_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_2.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void isRadioButtonAdv4_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_3.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

        #region ----- KeyDown Event -----

        private void PERSON_NAME_1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void PERSON_NAME_2_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void PERSON_NAME_3_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void DATE_FR_2_EditValueChanged(object pSender)
        {
            DATE_TO_2.EditValue = ISDate.ISMonth_Last(DATE_FR_2.DateTimeValue);
        }

        #endregion


    }
}