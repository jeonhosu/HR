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

namespace HRMF0402
{
    public partial class HRMF0402 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0402()
        {
            InitializeComponent();
        }

        public HRMF0402(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(INSUR_YYYYMM_0.EditValue) == string.Empty)
            {// 조회년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (itbINSUR.SelectedTab.TabIndex == 1)
            {
                igrHEALTH_INSUR.LastConfirmChanges();
                idaHEALTH_INSUR.OraSelectData.AcceptChanges();
                idaHEALTH_INSUR.Refillable = true;
                
                idaHEALTH_INSUR.Fill();
                igrHEALTH_INSUR.Focus();
            }
            else if (itbINSUR.SelectedTab.TabIndex == 2)
            {
                igrPENSION_INSUR.LastConfirmChanges();
                idaPENSION_INSUR.OraSelectData.AcceptChanges();
                idaPENSION_INSUR.Refillable = true;
                
                idaPENSION_INSUR.Fill();
                igrPENSION_INSUR.Focus();
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
            vSaveFileName = "Insurance List";     //기본 파일명. 수정필요.

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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_INSUR.Update();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_INSUR.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_INSUR.Cancel();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_INSUR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_INSUR.Delete();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_INSUR.Delete();
                    }
                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (idaHEALTH_INSUR.IsFocused)
                    {
                        ExcelExport(idaHEALTH_INSUR, igrHEALTH_INSUR); 
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        ExcelExport(idaPENSION_INSUR, igrPENSION_INSUR); 
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0402_Load(object sender, EventArgs e)
        {            
        }

        private void HRMF0402_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();                  // Corp Default Value Setting.

            INSUR_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            
            idaHEALTH_INSUR.FillSchema();
            idaPENSION_INSUR.FillSchema();

            INSUR_YYYYMM_0.Focus();
        }

        private void V_M_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int r = 0; r < igrHEALTH_INSUR.RowCount; r++)
            {
                igrHEALTH_INSUR.SetCellValue(r, igrHEALTH_INSUR.GetColumnToIndex("INSUR_YN"), V_M_SELECT_YN.CheckBoxString);
            }
        }

        private void V_P_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int r = 0; r < igrPENSION_INSUR.RowCount; r++)
            {
                igrPENSION_INSUR.SetCellValue(r, igrPENSION_INSUR.GetColumnToIndex("INSUR_YN"), V_P_SELECT_YN.CheckBoxString);
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string vYYYYMM = iConv.ISNull(INSUR_YYYYMM_0.EditValue);
            string vYYYY = vYYYYMM.Substring(0, 4);
            string vMM = vYYYYMM.Substring(5, 2);
            int vYYYY_Integer = int.Parse(vYYYY);
            int vMM_Integer = int.Parse(vMM);
            System.DateTime vDateTime = iDate.ISMonth_Last(new System.DateTime(vYYYY_Integer, vMM_Integer, 1));
            ildPERSON_0.SetLookupParamValue("W_WORK_DATE_TO", vDateTime);
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_EMPLYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----


        #endregion

    }
}