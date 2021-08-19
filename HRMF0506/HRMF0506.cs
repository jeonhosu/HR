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

namespace HRMF0506
{
    public partial class HRMF0506 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----
        public HRMF0506(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
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

        private void DefaultCorp()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            if (iString.ISNull(START_YYYYMM_0.EditValue) == String.Empty)
            {// 시작년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(END_YYYYMM_0.EditValue) == String.Empty)
            {// 종료년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_YYYYMM_0.Focus();
                return;
            }
            idaPAY_MASTER.SetSelectParamValue("W_STD_YYYYMM",END_YYYYMM_0.EditValue);
            idaPAY_MASTER.SetSelectParamValue("W_PAY_TYPE", PAY_GRADE_NAME_0.EditValue);

            idaPAY_MASTER.Fill();
            igrPERSON.Focus();
            
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

            object vDEPT_NAME = DEPT_NAME.EditValue;
            object vPAY_GRADE_NAME = PAY_GRADE_NAME.EditValue;
            object vPERSON_NAME = PERSON_NAME.EditValue;


            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = igrPAYMENT_TERM.RowCount;

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
                vSaveFileName = string.Format("개인별년급여조회_{0}~{1}", START_YYYYMM_0.EditValue, END_YYYYMM_0.EditValue);

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
                xlPrinting.OpenFileNameExcel = "HRMF0506_001.xlsx";
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
                    vPageNumber = xlPrinting.ExcelWrite(IDC_PRINTED_VALUE, igrPAYMENT_TERM, vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);

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

        private void XLPrinting_2(string pOutChoice)
        {// pOutChoice : 출력구분.
            //object mTitle = string.Empty;
            object mCORP_NAME = string.Empty;
            object mPERIOD_DATE = string.Empty;
            object mPRINTED_DATE = string.Empty;
            //object mPRINTED_BY = string.Empty;

            object vDEPT_NAME = DEPT_NAME.EditValue;
            object vPAY_GRADE_NAME = PAY_GRADE_NAME.EditValue;
            object vPERSON_NAME = PERSON_NAME.EditValue;


            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = igrPAYMENT_TERM.RowCount;

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
                vSaveFileName = string.Format("개인별년급여조회_{0}~{1}", START_YYYYMM_0.EditValue, END_YYYYMM_0.EditValue);

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
                xlPrinting.OpenFileNameExcel = "HRMF0506_002.xlsx";
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
                    vPageNumber = xlPrinting.ExcelWrite_2(IDC_PRINTED_VALUE, igrPAYMENT_TERM, vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);

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
                    if (idaPAYMENT_TERM.IsFocused)
                    {
                        idaPAYMENT_TERM.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaPAYMENT_TERM.IsFocused)
                    {
                        idaPAYMENT_TERM.AddUnder();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaPAY_MASTER.Update();                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPAY_MASTER.IsFocused)
                    {
                        idaPAY_MASTER.Cancel();
                    }
                    else if (idaPAYMENT_TERM.IsFocused)
                    {
                        idaPAYMENT_TERM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPAY_MASTER.IsFocused)
                    {
                        idaPAY_MASTER.Delete();
                    }
                    else if (idaPAYMENT_TERM.IsFocused)
                    {
                        idaPAYMENT_TERM.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_2("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_2("FILE");
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0506_Load(object sender, EventArgs e)
        {
            START_YYYYMM_0.EditValue = iDate.ISYear(DateTime.Today) + "-01".ToString();
            END_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
                        
            DefaultCorp();              //Default Corp.
        }
        #endregion  

        #region ----- LookUp Event -----

        private void ilaSTART_YYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
        }

        private void ilaEND_YYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", START_YYYYMM_0.EditValue);
        }

        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        
        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_START_DATE", iDate.ISMonth_1st(START_YYYYMM_0.EditValue));
            ildPERSON_0.SetLookupParamValue("W_END_DATE", iDate.ISMonth_Last(END_YYYYMM_0.EditValue));
        }

        #endregion

    }
}