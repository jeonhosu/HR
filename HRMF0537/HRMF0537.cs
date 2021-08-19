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

namespace HRMF0537
{
    public partial class HRMF0537 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0537()
        {
            InitializeComponent();
        }

        public HRMF0537(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            // 명세서 발급
            if (iString.ISNull(iedCORP_ID_0.EditValue) == string.Empty)
            {// 업체 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(iedPAY_YYYYMM.EditValue) == string.Empty)
            {// 지급일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedPAY_YYYYMM.Focus();
                return;
            }            
            if (iString.ISNull(iedWAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 선택 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedWAGE_TYPE_NAME.Focus();
                return;
            }
            
            // 그리드 부분 업데이트 처리
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            IDA_MONTH_PAYMENT.Fill();
            //idaALLOWANCE_INFO.Fill();
            //idaDEDUCTION_INFO.Fill();
            //idaDUTY_INFO.Fill();            
        }

        #endregion;

        // 인쇄 부분
        // Print 관련 소스 코드 2011.1.15(토)
        // Print 관련 소스 코드 2011.5.11(수) 수정
        #region ----- Convert String Method ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pCourse)
        {
            System.DateTime vStartTime = DateTime.Now;
            
            string vMessageText = string.Empty;

            string vBoxCheck = string.Empty;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            
            int vCountCheck = 0;
            object vObject = null;
            int vCountRow = IGR_MONTH_PAYMENT.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;

            int vIndexWAGE_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

            int vIndexPRINT_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("PRINT_TYPE"); 
            int vIndexPAY_YYYYMM  = IGR_MONTH_PAYMENT.GetColumnToIndex("PAY_YYYYMM"); 
            int vIndexPERSON_ID  = IGR_MONTH_PAYMENT.GetColumnToIndex("PERSON_ID");
            int vIndexNAME = IGR_MONTH_PAYMENT.GetColumnToIndex("NAME");
            int vIndexPERSON_NUM = IGR_MONTH_PAYMENT.GetColumnToIndex("PERSON_NUM");
            int vIndexCORP_ID  = IGR_MONTH_PAYMENT.GetColumnToIndex("CORP_ID");

            int vIndexCheckBox = IGR_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            string vCheckedString = IGR_MONTH_PAYMENT.GridAdvExColElement[vIndexCheckBox].CheckedString;
            //-------------------------------------------------------------------------------------
            for (int vRow = 0; vRow < vCountRow; vRow++)
            {
                vObject = IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexCheckBox);
                vBoxCheck = ConvertString(vObject);
                if (vBoxCheck == vCheckedString)
                {
                    vCountCheck++;
                }
            }

            if (vCountCheck < 1)
            {
                vMessageText = string.Format("Not Select");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }
            //-------------------------------------------------------------------------------------

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor; 
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0537_001.xls";
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(pCourse, IGR_MONTH_PAYMENT, IDA_PAY_ALLOWANCE, IDA_PAY_DEDUCTION, IDA_DUTY_INFO, IDA_BONUS_ALLOWANCE, IDA_BONUS_DEDUCTION);
            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessageText = string.Format("Printing End [Total Page : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();


            //이메일 발송 대상자 이메일 발송 처리 //
            vCountCheck = 0;
            for (int vRow = 0; vRow < IGR_MONTH_PAYMENT.RowCount; vRow++)
            {
                if (ConvertString(IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexPRINT_TYPE)) == "2")
                {
                    vCountCheck++;
                }
            }
            if (vCountCheck == 0)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
                return;
            }
            if (iString.ISNull(V_EMAIL_ACCOUNT_ID.EditValue) == string.Empty)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(string.Format("Sender : {0}", isMessageAdapter1.ReturnText("FCM_10256")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                V_EMAIL_ACCOUNT_ID.Focus();
                return;
            }
            if (iString.ISNull(V_EMAIL_ACCOUNT_PWD.EditValue) == string.Empty)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(string.Format("Sender : {0}", isMessageAdapter1.ReturnText("EAPP_10143")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                V_EMAIL_ACCOUNT_PWD.Focus();
                return;
            }

            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("E-Mail Send Data Save Start");
            System.Windows.Forms.Application.DoEvents();

            vCountCheck = 0;
            for (int vRow = 0; vRow < IGR_MONTH_PAYMENT.RowCount; vRow++)
            {
                if (ConvertString(IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexPRINT_TYPE)) == "2")
                {
                    vObject = IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexCheckBox);
                    vBoxCheck = ConvertString(vObject);                
                    if (vBoxCheck == vCheckedString)
                    {//체크한 대상중에 인쇄대상건만 인쇄//
                        vCountCheck++;

                        IGR_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIndexCheckBox);
                        IGR_MONTH_PAYMENT.Focus();
                        IGR_MONTH_PAYMENT.CurrentCellActivate(vRow, vIndexCheckBox);

                        vMessageText = string.Format("E-Mail Sending...{0}({1})", IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexNAME),
                                                                            IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexPERSON_NUM));
                        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                        System.Windows.Forms.Application.DoEvents();

                        IDC_SAVE_MONTH_PAYMENT_EMAIL.SetCommandParamValue("P_PAY_YYYYMM", IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexPAY_YYYYMM));
                        IDC_SAVE_MONTH_PAYMENT_EMAIL.SetCommandParamValue("P_WAGE_TYPE", IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexWAGE_TYPE));
                        IDC_SAVE_MONTH_PAYMENT_EMAIL.SetCommandParamValue("P_PERSON_ID", IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexPERSON_ID));
                        IDC_SAVE_MONTH_PAYMENT_EMAIL.SetCommandParamValue("P_CORP_ID", IGR_MONTH_PAYMENT.GetCellValue(vRow, vIndexCORP_ID));
                        IDC_SAVE_MONTH_PAYMENT_EMAIL.ExecuteNonQuery();
                        vSTATUS = iString.ISNull(IDC_SAVE_MONTH_PAYMENT_EMAIL.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_SAVE_MONTH_PAYMENT_EMAIL.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_SAVE_MONTH_PAYMENT_EMAIL.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            System.Windows.Forms.Cursor.Current = Cursors.Default;
                            System.Windows.Forms.Application.DoEvents();

                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        IGR_MONTH_PAYMENT.SetCellValue(vRow, vIndexCheckBox, "N");
                    }
                }
            }
            if (vCountCheck > 0)
            {//이메일 발송대상 존재=>이메일 발송처리
                vMessageText = string.Format("E-Mail Sending...Start");
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                IDC_SEND_EMAIL_MONTH_PAYMENT.ExecuteNonQuery();
                vSTATUS = iString.ISNull(IDC_SEND_EMAIL_MONTH_PAYMENT.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(IDC_SEND_EMAIL_MONTH_PAYMENT.GetCommandParamValue("O_MESSAGE"));
                if (IDC_SEND_EMAIL_MONTH_PAYMENT.ExcuteError || vSTATUS == "F")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();

                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                } 
            }
            vMessageText = string.Format("E-Mail Sending...Compeleted");
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
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
                    SearchDB();
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
                    XLPrinting_1("PRINT"); // 출력 함수 호출

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                    // 인쇄 완료 메시지 출력
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE"); // 출력 함수 호출
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0537_Load(object sender, EventArgs e)
        {                   
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            iedCORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            iedCORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            iedCORP_NAME_0.BringToFront();

            iedPAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            iedSTART_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            iedEND_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);

            // 그리드 부분 업데이트 처리 위함.
            IDA_MONTH_PAYMENT.FillSchema();

            //E MAIL전송자 정보
            IDC_GET_EMAIL_SENDER.ExecuteNonQuery();
            V_EMAIL_ACCOUNT_ID.EditValue = IDC_GET_EMAIL_SENDER.GetCommandParamValue("O_EMAIL_ACCOUNT_ID");
            V_EMAIL_ACCOUNT_PWD.EditValue = IDC_GET_EMAIL_SENDER.GetCommandParamValue("O_EMAIL_ACCOUNT_PWD");

            isAppInterfaceAdv1.OnAppMessage("");
        }

        // 전체선택 버튼
        private void btnSELECT_ALL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < IGR_MONTH_PAYMENT.RowCount; i++)
            {
                IGR_MONTH_PAYMENT.SetCellValue(i, IGR_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN"), "Y");
            }
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            string vMessageText = string.Format("Select Row [{0}]", IGR_MONTH_PAYMENT.RowCount);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();
        }

        // 취소 버튼
        private void btnCONFIRM_CANCEL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < IGR_MONTH_PAYMENT.RowCount; i++)
            {
                IGR_MONTH_PAYMENT.SetCellValue(i, IGR_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN"), "N");
            }            
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            isAppInterfaceAdv1.OnAppMessage("Select Row [0]");
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion

        #region ----- Lookup Event ----- 

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPRINT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PRINT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}