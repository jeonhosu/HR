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
using Syncfusion.XlsIO;

namespace HRMF0319
{
    public partial class HRMF0319 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0319(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            CORP_NAME_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                if (CORP_ID_0.EditValue == null)
                {// 업체.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CORP_NAME_0.Focus();
                    return;
                }
                if (DUTY_TYPE_0.EditValue != null && string.IsNullOrEmpty(DUTY_TYPE_0.EditValue.ToString()))
                {// 근무일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10059"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DUTY_TYPE_NAME_0.Focus();
                    return;
                }
                if (DUTY_YYYYMM_0.EditValue == null)
                {// 근무일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DUTY_YYYYMM_0.Focus();
                    return;
                }

                string vPERSON_NUM = iConv.ISNull(IGR_MONTH_TOTAL_SPREAD.GetCellValue("PERSON_NUM"));
                int vIDX_Col = IGR_MONTH_TOTAL_SPREAD.GetColumnToIndex("PERSON_NUM");

                IDA_MONTH_TOTAL_SPREAD.Fill();
                IGR_MONTH_TOTAL_SPREAD.Focus();

                if (IGR_MONTH_TOTAL_SPREAD.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_MONTH_TOTAL_SPREAD.RowCount; vRow++)
                    {
                        if (vPERSON_NUM == iConv.ISNull(IGR_MONTH_TOTAL_SPREAD.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_MONTH_TOTAL_SPREAD.CurrentCellMoveTo(vRow, vIDX_Col);
                        }
                    }
                }
            }

            if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                if (CORP_ID_1.EditValue == null)
                {// 업체.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CORP_NAME_1.Focus();
                    return;
                }
                if (DUTY_TYPE_1.EditValue != null && string.IsNullOrEmpty(DUTY_TYPE_0.EditValue.ToString()))
                {// 근무일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10059"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DUTY_TYPE_NAME_1.Focus();
                    return;
                }
                if (DUTY_YYYYMM_FR.EditValue == null)
                {// 근무일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DUTY_YYYYMM_FR.Focus();
                    return;
                }

                if (DUTY_YYYYMM_TO.EditValue == null)
                {// 근무일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DUTY_YYYYMM_TO.Focus();
                    return;
                }

                string vPERSON_NUM = iConv.ISNull(IGR_MONTH_PERIOD_SPREAD.GetCellValue("PERSON_NUM"));
                int vIDX_Col = IGR_MONTH_PERIOD_SPREAD.GetColumnToIndex("PERSON_NUM");

                IDA_MONTH_PERIOD_SPREAD.Fill();
                IGR_MONTH_PERIOD_SPREAD.Focus();

                if (IGR_MONTH_PERIOD_SPREAD.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_MONTH_PERIOD_SPREAD.RowCount; vRow++)
                    {
                        if (vPERSON_NUM == iConv.ISNull(IGR_MONTH_PERIOD_SPREAD.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_MONTH_PERIOD_SPREAD.CurrentCellMoveTo(vRow, vIDX_Col);
                        }
                    }
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
            vSaveFileName = "Monthly List";     //기본 파일명. 수정필요.

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
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(TB_BASE.SelectedTab.TabIndex == TP_MONTH.TabIndex)
                    {
                        ExcelExport(IDA_MONTH_TOTAL_SPREAD, IGR_MONTH_TOTAL_SPREAD);
                    }
                    else if(TB_BASE.SelectedTab.TabIndex == TP_PERIOD.TabIndex)
                    {
                        ExcelExport(IDA_MONTH_PERIOD_SPREAD, IGR_MONTH_PERIOD_SPREAD);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0319_Shown(object sender, EventArgs e)
        {
            // Year Month Setting
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
            DUTY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            WORK_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            WORK_DATE_TO.EditValue = iDate.ISMonth_Last(DateTime.Today);

            DUTY_YYYYMM_FR.EditValue = iDate.ISYearMonth(DateTime.Today, -1, 0);
            DUTY_YYYYMM_TO.EditValue = iDate.ISYearMonth(DateTime.Today);


            // CORP SETTING
            DefaultCorporation();

            // Duty TYPE SETTING
            ildDUTY_TYPE_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_MONTH_DUTY_TYPE.ExecuteNonQuery();
            DUTY_TYPE_NAME_0.EditValue = idcDEFAULT_MONTH_DUTY_TYPE.GetCommandParamValue("O_CODE_NAME");
            DUTY_TYPE_0.EditValue = idcDEFAULT_MONTH_DUTY_TYPE.GetCommandParamValue("O_CODE");

            DUTY_TYPE_NAME_1.EditValue = idcDEFAULT_MONTH_DUTY_TYPE.GetCommandParamValue("O_CODE_NAME");
            DUTY_TYPE_1.EditValue = idcDEFAULT_MONTH_DUTY_TYPE.GetCommandParamValue("O_CODE");

            // LEAVE CLOSE TYPE SETTING
            ildCLOSE_FLAG_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ildCLOSE_FLAG_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            CLOSED_FLAG_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            CLOSED_FLAG_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            CLOSED_FLAG_NAME_1.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            CLOSED_FLAG_1.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            WORK_DATE_FR.BringToFront();
            WORK_DATE_TO.BringToFront();

            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
        }

        #endregion  

        #region ----- Adapter Event -----

        #endregion

        #region ----- LookUp Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDUTY_TYPE_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaDUTY_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDUTY_TYPE_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion


        

       


    }
}