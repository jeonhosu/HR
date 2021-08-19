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

namespace HRMF0326
{
    public partial class HRMF0326 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0326(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

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

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_START_DATE.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_START_DATE.Focus();
                return;
            }
                        
            IDA_PERSON_WORK_PERIOD.Fill();
            IGR_PERSON_WORK_PERIOD.Focus();
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

            if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                vExport.GridToExcel(vGrid.BaseGrid, vSaveFileDialog.FileName,
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
        }

        #endregion


        #region ----- XL Print Method -----

        private void XLPrinting(string pOutChoice)
        {
            //print type 설정
            DialogResult vdlgResult;
            HRMF0326_PRINT vHRMF0326_PRINT = new HRMF0326_PRINT(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0326_PRINT.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }

            string vPRINT_TYPE = iString.ISNull(vHRMF0326_PRINT.Get_Printer_Type);
            if (iString.ISNull(vPRINT_TYPE) == string.Empty)
            {
                return;
            }
            string vPRINT_PREVIEW_YN = iString.ISNull(vHRMF0326_PRINT.Get_Print_Preview);

            //부서별
            XLPrinting_1(vPRINT_TYPE, vPRINT_PREVIEW_YN); 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void XLPrinting_1(string pOutChoice, string pPRINT_PREVIEW_YN)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;
 
            int vCountRow = IGR_PERSON_WORK_PERIOD.RowCount;
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            if (pOutChoice == "PDF")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "pdf file(*.pdf)|*.pdf";
                vSaveFileDialog.DefaultExt = "pdf";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }

            vMessageText = string.Format(" Printing Starting..."); 
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                // open해야 할 파일명 지정.
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0326_001.xlsx";
                //-------------------------------------------------------------------------------------
                //// 파일 오픈.
                ////-------------------------------------------------------------------------------------
                //bool isOpen = xlPrinting.XLFileOpen();
                ////-------------------------------------------------------------------------------------

                ////-------------------------------------------------------------------------------------
                //if (isOpen == true)
                //{
                //    //헤더 데이터 설정
                string vPERIOD_NAME = string.Format("({0} ~ {1})", iDate.ISGetDate(W_START_DATE.EditValue).ToShortDateString(),
                                                                    iDate.ISGetDate(W_END_DATE.EditValue).ToShortDateString());

                //    //헤더 인쇄
                //    xlPrinting.HeaderWrite_1(W_CORP_NAME.EditValue, vPERIOD_NAME);

                    //라인 인쇄
                    vPageNumber = xlPrinting.LineWrite_1(W_CORP_NAME.EditValue, vPERIOD_NAME, IDA_PERSON_WORK_PERIOD);

                    //출력구분에 따른 선택(인쇄 or file 저장)
                    if (pOutChoice == "PRINTER")
                    {
                        if (pPRINT_PREVIEW_YN == "Y")
                        {
                            xlPrinting.Preview_Printing(1, vPageNumber);
                        }
                        else
                        {
                            xlPrinting.Printing(1, vPageNumber);
                        }
                    }
                    else if (pOutChoice == "PDF")
                    {
                        xlPrinting.PDF(vSaveFileName);
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                //}
                //else
                //{
                //    vMessageText = "Excel File Open Error";
                //    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                //    System.Windows.Forms.Application.DoEvents();
                //}
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

        //private void XLPrinting_2(string pOutChoice)
        //{
        //    //예산신청내역 - 계정별
        //    string vMessageText = string.Empty;
        //    string vSaveFileName = string.Empty;

        //    IDA_PRINT_BUDGET_ACCOUNT.Fill();
        //    int vCountRow = IDA_PRINT_BUDGET_ACCOUNT.OraSelectData.Rows.Count;
        //    if (vCountRow < 1)
        //    {
        //        vMessageText = string.Format("Without Data");
        //        isAppInterfaceAdv1.OnAppMessage(vMessageText);
        //        System.Windows.Forms.Application.DoEvents();
        //        return;
        //    }

        //    //출력구분이 파일인 경우 처리.
        //    if (pOutChoice == "FILE")
        //    {
        //        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        //        vSaveFileName = "Budget_assign_account";

        //        saveFileDialog1.Title = "Excel Save";
        //        saveFileDialog1.FileName = vSaveFileName;
        //        saveFileDialog1.Filter = "Excel file(*.xls)|*.xls";
        //        saveFileDialog1.DefaultExt = "xls";
        //        if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        //        {
        //            return;
        //        }
        //        else
        //        {
        //            vSaveFileName = saveFileDialog1.FileName;
        //            System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
        //            try
        //            {
        //                if (vFileName.Exists)
        //                {
        //                    vFileName.Delete();
        //                }
        //            }
        //            catch (Exception EX)
        //            {
        //                MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                return;
        //            }
        //        }
        //        vMessageText = string.Format(" Writing Starting...");
        //    }
        //    else
        //    {
        //        vMessageText = string.Format(" Printing Starting...");
        //    }

        //    System.Windows.Forms.Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Application.DoEvents();

        //    int vPageNumber = 0;
        //    XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

        //    try
        //    {
        //        // open해야 할 파일명 지정.
        //        //-------------------------------------------------------------------------------------
        //        xlPrinting.OpenFileNameExcel = "FCMF0621_002.xls";
        //        //-------------------------------------------------------------------------------------
        //        // 파일 오픈.
        //        //-------------------------------------------------------------------------------------
        //        bool isOpen = xlPrinting.XLFileOpen();
        //        //-------------------------------------------------------------------------------------

        //        //-------------------------------------------------------------------------------------
        //        if (isOpen == true)
        //        {
        //            //헤더 데이터 설정
        //            object vBUDGET_YEAR = BUDGET_YEAR_0.EditValue;

        //            //헤더 인쇄
        //            xlPrinting.HeaderWrite_2(vBUDGET_YEAR);
        //            //라인 인쇄
        //            vPageNumber = xlPrinting.LineWrite_2(IDA_PRINT_BUDGET_ACCOUNT);

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

        //            vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
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
                    if (IDA_PERSON_WORK_PERIOD.IsFocused)
                    {
                        IDA_PERSON_WORK_PERIOD.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExcelExport(IGR_PERSON_WORK_PERIOD);
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0326_Load(object sender, EventArgs e)
        {
            this.Visible = true;

            IDA_PERSON_WORK_PERIOD.FillSchema();
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);
            
            // CORP SETTING
            DefaultCorporation();

            // LEAVE CLOSE TYPE SETTING
            ILD_CLOSE_FLAG_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ILD_CLOSE_FLAG_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            W_CLOSE_FLAG_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            W_CLOSE_FLAG.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            W_CORP_NAME.BringToFront();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
        }

        private void START_DATE_0_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            W_END_DATE.EditValue = iDate.ISMonth_Last(e.EditValue);
        }
              
        #endregion  

        #region ----- Adapter Event -----
         
        #endregion

        #region ----- LookUp Event -----

        private void ilaOPERATING_UNIT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", W_CORP_ID.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

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

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE",W_END_DATE.EditValue);
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

        private void ilaHOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}