using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using System.IO;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using System.IO;
using Syncfusion.XlsIO;

namespace HRMF0223
{
    public partial class HRMF0223 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0223()
        {
            InitializeComponent();
        }

        public HRMF0223(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP_1.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP_1.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_2.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_2.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            CORP_NAME_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (TB_EDUCATION.SelectedTab.TabIndex == 2)
            {
                if (iConv.ISNull(CORP_ID_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true; 
                    return;
                }

                if (iConv.ISNull(STAT_DATE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                if (iConv.ISNull(END_DATE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                if (Convert.ToDateTime(STAT_DATE.EditValue) > Convert.ToDateTime(END_DATE.EditValue))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    STAT_DATE.Focus();
                    return;
                }

                IDA_EDUCATION_STATE_1.Fill();
                IGR_EDUCATION_CURRENT.Focus(); 
            }

            else if (TB_EDUCATION.SelectedTab.TabIndex == 3 )
            {
                

                if (iConv.ISNull(CORP_ID_2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                if (iConv.ISNull(STD_DATE_2.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                IDA_PERSON_2.Fill();
                IGR_PERSON_INFO.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
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

        #region ----- Excel Export II -----

        private void ExcelExport(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.RestoreDirectory = true;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Education List";     //기본 파일명. 수정필요.

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
                    if (IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        IDA_EDUCATION_STATE_1.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        IDA_EDUCATION_STATE_1.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        IDA_EDUCATION_STATE_1.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        IDA_EDUCATION_STATE_1.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        IDA_EDUCATION_STATE_1.Delete();
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(IDA_EDUCATION_STATE_1.IsFocused)
                    {
                        ExcelExport(IDA_EDUCATION_STATE_1, IGR_EDUCATION_CURRENT);
                    }
                    else if(IDA_EDUCATION_STATE_2.IsFocused)
                    {
                        ExcelExport(IDA_EDUCATION_STATE_2, IGR_EDUCATION_STATE_2);
                    }
                }
            }
        }

        #endregion;

        #region ----Form event -----

        private void HRMF0223_Load(object sender, EventArgs e)
        {
            IDA_EDUCATION_STATE_1.FillSchema();
        }

        private void HRMF0223_Shown(object sender, EventArgs e)
        {
            DefaultValues();

            STAT_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = DateTime.Today;
            STD_DATE_2.EditValue = DateTime.Today;

            CORP_NAME_1.BringToFront();
            CORP_NAME_2.BringToFront();
        }

        private void IGR_EDUCATION_CURRENT_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_START_DATE = IGR_EDUCATION_CURRENT.GetColumnToIndex("START_DATE");
            int vIDX_END_DATE = IGR_EDUCATION_CURRENT.GetColumnToIndex("END_DATE");
            object vSTART_DATE;
            object vEND_DATE;
            if (vIDX_START_DATE == e.ColIndex)
            {
                vSTART_DATE = e.NewValue;
            }
            else
            {
                vSTART_DATE = IGR_EDUCATION_CURRENT.GetCellValue("START_DATE");
            }

            if (vIDX_END_DATE == e.ColIndex)
            {
                vEND_DATE = e.NewValue;
            }
            else
            {
                vEND_DATE = IGR_EDUCATION_CURRENT.GetCellValue("END_DATE");
            }

            IDC_EDU_GET_TIME_P.SetCommandParamValue("P_START_DATE", vSTART_DATE);
            IDC_EDU_GET_TIME_P.SetCommandParamValue("P_END_DATE", vEND_DATE);
            IDC_EDU_GET_TIME_P.ExecuteNonQuery();
            IGR_EDUCATION_CURRENT.SetCellValue("EDU_TIME", IDC_EDU_GET_TIME_P.GetCommandParamValue("O_EDU_TIME"));
        }

        #endregion

        #region ------ Lookup event ------

        private void ILA_PERSON_0_SelectedRowData(object pSender)
        {
            Search_DB();
        }

        private void ILA_PERSON_1_SelectedRowData(object pSender)
        {
            Search_DB();
        }

        private void ILA_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_2.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_DEPT_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_1.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_FLOOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_POST_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        #endregion 


        #region ------ button event ------

        private void bXL_Choice_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_EXCEL_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0223_EXPORT vHRMF0223_EXPORT = new HRMF0223_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , CORP_ID_1.EditValue, CORP_NAME_1.EditValue);
            vdlgResult = vHRMF0223_EXPORT.ShowDialog(); 
            if (vdlgResult == DialogResult.OK)
            {
                 
            }
            vHRMF0223_EXPORT.Dispose();
        }

        private void bXL_UpLoad_Deduction_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //Excel_Upload();  
            DialogResult vdlgResult;
            HRMF0223_UPLOAD vHHRMF0223_UPLOAD = new HRMF0223_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface, CORP_ID_1.EditValue);
            vdlgResult = vHHRMF0223_UPLOAD.ShowDialog();
            vHHRMF0223_UPLOAD.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        #endregion

        #region ----- Excel Upload : Asset Master -----

        private void Select_Excel_File()
        {
            //try
            //{
            //    DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

            //    openFileDialog1.Title = "Select Open File";
            //    openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
            //    openFileDialog1.DefaultExt = "xls";
            //    openFileDialog1.FileName = "*.xls;*.xlsx";
            //    if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //    {
            //        FILE_PATH_MASTER.EditValue = openFileDialog1.FileName;
            //    }
            //    else
            //    {
            //        FILE_PATH_MASTER.EditValue = string.Empty;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    isAppInterfaceAdv1.OnAppMessage(ex.Message);
            //    Application.DoEvents();
            //}
        }

        private void Excel_Upload()
        {
            //string vSTATUS = string.Empty;
            //string vMESSAGE = string.Empty;
            //bool vXL_Load_OK = false;

            //if (iConv .ISNull(FILE_PATH_MASTER.EditValue) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH_MASTER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            //Application.UseWaitCursor = true;
            //this.Cursor = Cursors.WaitCursor;
            //Application.DoEvents();

            //string vOPenFileName = FILE_PATH_MASTER.EditValue.ToString();
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
            //    return;
            //}


            //// 업로드 아답터 fill //
            //IDA_XLUPLOAD_EDUCATION .Fill();

            //try
            //{
            //    if (vXL_Load_OK == true)
            //    {
            //        vXL_Load_OK = vXL_Upload.LoadXL_10(IDA_XLUPLOAD_EDUCATION, 2);

            //        if (vXL_Load_OK == false)
            //        {
            //            IDA_XLUPLOAD_EDUCATION.Cancel();
            //            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        }
            //        else
            //        {
            //            IDA_XLUPLOAD_EDUCATION.Update();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    IDA_XLUPLOAD_EDUCATION.Cancel();

            //    isAppInterfaceAdv1.OnAppMessage(ex.Message);

            //    vXL_Upload.DisposeXL();

            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();
            //    return;
            //}
            //vXL_Upload.DisposeXL();

            //Application.UseWaitCursor = false;
            //this.Cursor = Cursors.Default;
            //Application.DoEvents();
        }



        #endregion

    }
}