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

namespace HRMF0516
{
    public partial class HRMF0516 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #endregion;

        #region ----- Constructor -----

        public HRMF0516()
        {
            InitializeComponent();
        }

        public HRMF0516(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iConv.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                IDA_SALARY_ITEM_SUM.Fill();
                IGR_SALARY_ITEM_SUM.Focus();
            }
            else
            {   
                IDA_SALARY_ALLOWANCE_SUM.Fill();
                IGR_SALARY_ALLOWANCE_SUM.Focus();
            }
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
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

        #region ----- XL Print Method -----

        private void XLPrinting_Main(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object vPay_YYYYMM = PAY_YYYYMM_0.EditValue;
            object vWAGE_TYPE_NAME = WAGE_TYPE_NAME_0.EditValue;
            object vWAGE_TYPE = WAGE_TYPE_0.EditValue;

            //object mPRINTED_BY = string.Empty;

            string vMessageText = string.Empty;
            string vSaveFileName = "Salary_Item_";
            
            //Data 체크
            int vCountRow = 0;
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                //지급 항목별 상세
                vCountRow = IGR_SALARY_ALLOWANCE_SUM.RowCount;
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                //지급/공제 항목별 상세
                vCountRow = IGR_SALARY_ITEM_SUM.RowCount;
            }
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("{0}{1}{2}", vSaveFileName, WAGE_TYPE_NAME_0.EditValue, PAY_YYYYMM_0.EditValue);

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

            //tab에 따른 인쇄 선택
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                //지급 항목별 상세
                XLPrinting_1(pOutChoice, vSaveFileName, vPay_YYYYMM, vWAGE_TYPE_NAME, vWAGE_TYPE);
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                //지급/공제 항목별 상세
                XLPrinting_2(pOutChoice, vSaveFileName, vPay_YYYYMM, vWAGE_TYPE_NAME, vWAGE_TYPE);
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        //지급 항목별 상세
        private void XLPrinting_1(string pOutChoice, string pSaveFileName, object pPay_YYYYMM, object pWAGE_TYPE_NAME, object pWAGE_TYPE)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;

            int vCountRow = IGR_SALARY_ALLOWANCE_SUM.RowCount;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                if (pSaveFileName == string.Empty)
                {
                    MessageBoxAdv.Show("FileName is not entered", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
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
                xlPrinting.OpenFileNameExcel = "HRMF0516_001.xls";
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
                    vPageNumber = xlPrinting.ExcelWrite_1(IDC_PRINTED_VALUE, IGR_SALARY_ALLOWANCE_SUM, pPay_YYYYMM, pWAGE_TYPE_NAME, pWAGE_TYPE);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(pSaveFileName);
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

        //지급/공제 항목 세부내역
        private void XLPrinting_2(string pOutChoice, string pSaveFileName, object pPay_YYYYMM, object pWAGE_TYPE_NAME, object pWAGE_TYPE)
        {// pOutChoice : 출력구분.
            string vMessageText = string.Empty;

            int vCountRow = IGR_SALARY_ITEM_SUM.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            if (pOutChoice == "FILE")
            {
                if (pSaveFileName == string.Empty)
                {
                    MessageBoxAdv.Show("FileName is not entered", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
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
                xlPrinting.OpenFileNameExcel = "HRMF0516_002.xls";
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
                    vPageNumber = xlPrinting.ExcelWrite_2(IDC_PRINTED_VALUE, IGR_SALARY_ITEM_SUM, pPay_YYYYMM, pWAGE_TYPE_NAME, pWAGE_TYPE);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE(pSaveFileName);
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
                    if (IDA_SALARY_ITEM_SUM.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_SALARY_ITEM_SUM.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_SALARY_ITEM_SUM.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SALARY_ITEM_SUM.IsFocused)
                    {
                        IDA_SALARY_ITEM_SUM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_SALARY_ITEM_SUM.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0516_Load(object sender, EventArgs e)
        {

        }

        private void HRMF0516_Shown(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);

            DefaultCorporation();              //Default Corp.
        }

        #endregion

        #region ----- Lookup Event ------

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ILA_JOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("JOB_CATEGORY", "Y");
        }

        #endregion

    }
}