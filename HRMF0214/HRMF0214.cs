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

namespace HRMF0214
{
    public partial class HRMF0214 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0214()
        {
            InitializeComponent();
        }

        public HRMF0214(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_1.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_1.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                if (iConv.ISNull(STD_DATE_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //e.Cancel = true;
                    return;
                }

                IDA_ENROLL_ADMINISTRATIVE_LEAVE.Fill();
                IGR_ENROLL_ADMINISTRATIVE_LEAVE.Focus();
            }

            if (TB_BASE.SelectedTab.TabIndex == 2)
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

                IDA_INQUIRY_ADMINISTRATIVE_LEAVE.Fill();
                IGR_INQUIRY_ADMINISTRATIVE_LEAVE.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object mCORP_NAME = string.Empty;
            object mPERIOD_DATE = string.Empty;
            object mPRINTED_DATE = string.Empty;
            //object mPRINTED_BY = string.Empty;

            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = IGR_INQUIRY_ADMINISTRATIVE_LEAVE.RowCount;

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
                vSaveFileName = string.Format("Leave_{0}",DateTime.Today.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
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
                xlPrinting.OpenFileNameExcel = "HRMF0214_001.xlsx";
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
                    vPageNumber = xlPrinting.ExcelWrite (IDC_PRINTED_VALUE, IGR_INQUIRY_ADMINISTRATIVE_LEAVE);

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
                    if (IDA_ENROLL_ADMINISTRATIVE_LEAVE.IsFocused)
                    {
                        IDA_ENROLL_ADMINISTRATIVE_LEAVE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ENROLL_ADMINISTRATIVE_LEAVE.IsFocused)
                    {
                        IDA_ENROLL_ADMINISTRATIVE_LEAVE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IGR_ENROLL_ADMINISTRATIVE_LEAVE.CurrentCellMoveTo(1, IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetColumnToIndex("PERSON_NUM"));
                    IDA_ENROLL_ADMINISTRATIVE_LEAVE.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ENROLL_ADMINISTRATIVE_LEAVE.IsFocused)
                    {
                        IDA_ENROLL_ADMINISTRATIVE_LEAVE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ENROLL_ADMINISTRATIVE_LEAVE.IsFocused)
                    {
                        IDA_ENROLL_ADMINISTRATIVE_LEAVE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form event ------

        private void HRMF0214_Load(object sender, EventArgs e)
        {
            IDA_ENROLL_ADMINISTRATIVE_LEAVE.FillSchema();
        }

        private void HRMF0214_Shown(object sender, EventArgs e)
        {
            DefaultValues();

            STD_DATE_0.EditValue = DateTime.Today;
            CB_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            STAT_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = DateTime.Today;
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

        private void ILA_PERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_0.SetLookupParamValue("W_CORP_ID", CORP_ID_0.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_DEPT_ID", DEPT_ID_0.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_FLOOR_ID", FLOOR_ID_0.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_START_DATE", STD_DATE_0.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
        }

        private void ILA_PERSON_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_0.SetLookupParamValue("W_CORP_ID", CORP_ID_1.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_DEPT_ID", DEPT_ID_1.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_FLOOR_ID", FLOOR_ID_1.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_START_DATE", STAT_DATE.EditValue);
            ILD_PERSON_0.SetLookupParamValue("W_END_DATE", END_DATE.EditValue);
        }

        private void ILA_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_DUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DUTY", "Y");
        }

        private void ILA_YYYYMM_TO_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            string vSTART_YYYYMM = "2010-01";
            string vEND_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("START_DATE")), 1));
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", vSTART_YYYYMM);
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", vEND_YYYYMM); 
        }

        private void ILA_YYYYMM_FR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            string vSTART_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("END_DATE")), -2));
            string vEND_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("END_DATE")), 6));
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", vSTART_YYYYMM);
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", vEND_YYYYMM); 
        }

        private void ILA_OFFICE_TAX_PERIOD_FR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            string vSTART_YYYYMM = "2010-01";
            string vEND_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("START_DATE")), -1));
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", vSTART_YYYYMM);
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", vEND_YYYYMM);
        }

        private void ILA_OFFICE_TAX_PERIOD_TO_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            string vSTART_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("END_DATE")), -2));
            string vEND_YYYYMM = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(IGR_ENROLL_ADMINISTRATIVE_LEAVE.GetCellValue("END_DATE")), 12));  
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", vSTART_YYYYMM);
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", vEND_YYYYMM);
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_FLOOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_JOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        private void ILA_JOB_CATEGORY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        #endregion
        
        #region ------ Adapter event ------

        private void IDA_ADMINISTRATIVE_LEAVE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["START_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["END_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["START_DATE"]) > iDate.ISGetDate(e.Row["END_DATE"]))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10473"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}