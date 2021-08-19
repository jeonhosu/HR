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

namespace HRMF0716
{
    public partial class HRMF0716 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0716()
        {
            InitializeComponent();
        }

        public HRMF0716(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iString.ISNull(W_CORP_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            IDA_YEAR_ADJUSTMENT_SPREAD.Fill();
            IGR_YEAR_ADJUSTMENT_SPREAD.Focus();

        }


        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }


        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport()
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog1.Title = "Save File Name";
            saveFileDialog1.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog1.DefaultExt = ".xls";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                vExport.GridToExcel(IGR_YEAR_ADJUSTMENT_SPREAD.BaseGrid, saveFileDialog1.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.RowHeaders);

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog1.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion


        #region ----- XL Print 1 Method -----

        //private void XLPrinting_1(string pOutChoice)
        //{
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;
        //    object vTerritory = string.Empty;

        //    object vSTD_YYYYMM = W_STD_YYYYMM.EditValue;

        //    int vCountRow = IGR_YEAR_ADJUSTMENT_SPREAD.RowCount;
        //    if (vCountRow < 1)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //    vSaveFileName = String.Format("연말정산내역_{0}", vSTD_YYYYMM);

        //    saveFileDialog1.Title = "Excel Save";
        //    saveFileDialog1.FileName = vSaveFileName;
        //    saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //    saveFileDialog1.DefaultExt = "xls";
        //    if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //    {
        //        return;
        //    }
        //    else
        //    {
        //        vSaveFileName = saveFileDialog1.FileName;
        //        System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //        if (vFileName.Exists)
        //        {
        //            try
        //            {
        //                vFileName.Delete();
        //            }
        //            catch (Exception EX)
        //            {
        //                MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                return;
        //            }
        //        }
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Application.DoEvents();

        //    int vPageNumber = 0;

        //    vMessageText = string.Format(" Printing Starting...");
        //    isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //    System.Windows.Forms.Application.DoEvents();

        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {// 폼에 있는 항목들중 기본적으로 출력해야 하는 값.

        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "HRMF0716_001.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            xlPrinting.HeaderWrite(IGR_YEAR_ADJUSTMENT_SPREAD, vSTD_YYYYMM);
        //            // 실제 인쇄
        //            vPageNumber = xlPrinting.LineWrite(IGR_YEAR_ADJUSTMENT_SPREAD);

        //            //출력구분에 따른 선택(인쇄 or file 저장)
        //            if (pOutChoice == "PRINT")
        //            {
        //                xlPrinting.Printing(1, vPageNumber);
        //            }
        //            else if (pOutChoice == "FILE")
        //            {
        //                xlPrinting.SAVE(vSaveFileName);
        //            }

        //            //-------------------------------------------------------------------------------------
        //            xlPrinting.Dispose();
        //            //-------------------------------------------------------------------------------------

        //            vMessageText = "Printing End";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        else
        //        {
        //            vMessageText = "Excel File Open Error";
        //            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //        //-------------------------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        xlPrinting.Dispose();

        //        vMessageText = ex.Message;
        //        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Application.DoEvents();
        //}

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
                    if (IDA_YEAR_ADJUSTMENT_SPREAD.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT_SPREAD.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_YEAR_ADJUSTMENT_SPREAD.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT_SPREAD.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_YEAR_ADJUSTMENT_SPREAD.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT_SPREAD.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_YEAR_ADJUSTMENT_SPREAD.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT_SPREAD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_YEAR_ADJUSTMENT_SPREAD.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT_SPREAD.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //XLPrinting_1("FILE");
                    ExcelExport();
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0716_Load(object sender, EventArgs e)
        {
            W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 4)));

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();

            // Standard Date SETTING
            //if (DateTime.Today.Month <= 2)
            //{
            //    DateTime dLastYearMonthDay = new DateTime(DateTime.Today.AddYears(-1).Year, 12, 31);
            //    STD_YYYYMM.EditValue = dLastYearMonthDay;
            //}
            //else
            //{
            //    DateTime dLastYearMonthDay = new DateTime(DateTime.Today.Year, 12, 31);
            //    STANDARD_DATE_0.EditValue = dLastYearMonthDay;
            //}            
            
            IDA_YEAR_ADJUSTMENT_SPREAD.FillSchema();
        }

        private void HRMF0716_Shown(object sender, EventArgs e)
        {            
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ILA_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaPOST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("POST", "Y");
        }

        private void IlaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("YEAR_EMPLOYE_TYPE", "Y");
        }


        #endregion

    }
}