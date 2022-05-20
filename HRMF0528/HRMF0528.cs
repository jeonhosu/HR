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

using System.Net.Mail;
using System.Net;
using System.Net.Sockets;

namespace HRMF0528
{
    public partial class HRMF0528 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0528()
        {
            InitializeComponent();
        }

        public HRMF0528(Form pMainForm, ISAppInterface pAppInterface)
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
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {// 업체 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PAY_YYYYMM.EditValue) == string.Empty)
            {// 지급일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PAY_YYYYMM.Focus();
                return;
            }            
            if (iConv.ISNull(W_WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 선택 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_PAPER.TabIndex)
            {
                // 그리드 부분 업데이트 처리
                IGR_MONTH_PAYMENT.LastConfirmChanges();
                IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
                IDA_MONTH_PAYMENT.Refillable = true;

                IDA_MONTH_PAYMENT.Fill();          
            }
            else if(TB_MAIN.SelectedTab.TabIndex == TP_EMAIL.TabIndex)
            {
                IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
                IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
                IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

                IDA_MONTH_PAYMENT_EMAIL.Fill();
            }
        }

        private void SUB_STATUS(bool pEnabled_Flag, string pSub_Form)
        {
            if(pEnabled_Flag == true)
            {
                if (pSub_Form == "EMAIL_TEXT")
                {
                    GB_EMAIL_TEXT.Left = 300;
                    GB_EMAIL_TEXT.Top = 120;

                    GB_EMAIL_TEXT.Width = 610;
                    GB_EMAIL_TEXT.Height = 360; 

                    GB_CONDITION.Enabled = false;
                    TB_MAIN.Enabled = true; //false;

                    GB_EMAIL_TEXT.BringToFront();
                    GB_EMAIL_TEXT.Visible = true;
                    GB_EMAIL_TEXT.Enabled = true; 
                    V_SAVE.Enabled = true;
                    V_CLOSE.Enabled = true;


                }
                else
                {
                    GB_CONDITION.Enabled = false;
                    TB_MAIN.Enabled = false;

                    GB_STATUS.Enabled = true;
                    GB_STATUS.BringToFront();
                    GB_STATUS.Visible = true;
                } 
            }
            else
            {
                V_PRINT.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                V_EMAIL_SEND.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                V_SAVE_FILE.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

                GB_CONDITION.Enabled = true;
                TB_MAIN.Enabled = true;

                GB_STATUS.BringToFront();
                GB_STATUS.Visible = false; 
                GB_EMAIL_TEXT.Visible = false;
            }
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
                xlPrinting.OpenFileNameExcel = "HRMF0528_001.xlsx";
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(pCourse, IGR_MONTH_PAYMENT, IDA_PAY_ALLOWANCE, IDA_PAY_DEDUCTION, IDA_MONTH_DUTY, IDA_BONUS_ALLOWANCE, IDA_BONUS_DEDUCTION, CB_STAMP.CheckBoxString);
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
             
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_Main(string pOutput_Type)
        {
            string vSaveFileName = string.Empty;
            if (pOutput_Type == "EXCEL")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "xlsx file(*.xlsx)|*.xlsx";
                vSaveFileDialog.DefaultExt = "xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }
            else if (pOutput_Type == "PDF")
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
                    return;
                }
            }
            

          //  IDC_GET_REPORT_SET.SetCommandParamValue("P_STD_DATE", GL_DATE.EditValue);
            IDC_GET_REPORT_SET.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0528");
            IDC_GET_REPORT_SET.ExecuteNonQuery();
            string vREPORT_TYPE = iConv.ISNull(IDC_GET_REPORT_SET.GetCommandParamValue("O_REPORT_TYPE"));
            string vREPORT_FILE_NAME = iConv.ISNull(IDC_GET_REPORT_SET.GetCommandParamValue("O_REPORT_FILE_NAME"));

            if (vREPORT_TYPE.ToUpper() == "SIK")
            {
                XLPrinting_SIK( pOutput_Type, vREPORT_FILE_NAME, CB_STAMP.CheckBoxString);
            }
            else if (vREPORT_TYPE.ToUpper() == "SIV" )
            {
                XLPrinting_SIV(vREPORT_FILE_NAME, pOutput_Type, CB_STAMP.CheckBoxString);
            }    
            else
            {
                XLPrinting_BSK(pOutput_Type, vREPORT_FILE_NAME, CB_STAMP.CheckBoxString);
            }          
        }

        private void XLPrinting_SIV(string pReport_File_Name, string pCourse, string pCB_STAMP)
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

            int vIndexWAGE_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");


            //-------------------------------------------------------------------------------------
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor; 
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                if (pReport_File_Name == string.Empty)
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0528_031.xlsx";
                }
                else
                {
                    xlPrinting.OpenFileNameExcel = pReport_File_Name;
                }

                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain_SIV(pCourse, IGR_MONTH_PAYMENT, IDA_PAY_ALLOWANCE, IDA_PAY_DEDUCTION, IDA_MONTH_DUTY, IDA_MONTH_OT);
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
             
            System.Windows.Forms.Cursor.Current = Cursors.Default; 
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_BSK(string pOUTPUT_TYPE, string pReport_File_Name, string pCB_STAMP)
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

            int vIndexWAGE_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

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

            IGR_MONTH_PAYMENT.LastConfirmChanges();
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
            IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            GB_CONDITION.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            
            try
            {
                //-------------------------------------------------------------------------------------
                if (pReport_File_Name == string.Empty)
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0528_003.xlsx";
                }
                else
                {
                    xlPrinting.OpenFileNameExcel = pReport_File_Name;
                }

                
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(pOUTPUT_TYPE, IGR_MONTH_PAYMENT, IDA_PAY_ALLOWANCE, IDA_PAY_DEDUCTION, IDA_MONTH_DUTY, IDA_MONTH_OT, CB_STAMP.CheckBoxString);
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

            this.Cursor = System.Windows.Forms.Cursors.Default;
            GB_CONDITION.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void XLPrinting_SIK(string pOUTPUT_TYPE, string pReport_File_Name, string pCB_STAMP)
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

            int vIndexWAGE_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = IGR_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

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

            IGR_MONTH_PAYMENT.LastConfirmChanges();
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
            IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            GB_CONDITION.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                if (pReport_File_Name == string.Empty)
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0528_002.xlsx";
                }
                else
                {
                    xlPrinting.OpenFileNameExcel = pReport_File_Name;
                }


                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain_SIK(pOUTPUT_TYPE, IGR_MONTH_PAYMENT, IDA_PAY_ALLOWANCE, IDA_PAY_DEDUCTION, IDA_MONTH_DUTY, IDA_MONTH_OT, pCB_STAMP);
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

            this.Cursor = System.Windows.Forms.Cursors.Default;
            GB_CONDITION.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- C# Email 발송 -----

        private void Send_eMail(string pRept_ID, string pName, string pAttachment_Path)
        {
            MailMessage vMail = new MailMessage();

            //보내는 사람 메일주소.
            vMail.From = new MailAddress(iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue));

            //받는 사람 메일주소.여러 사람에게 보낼경우 계속 추가하면 됨.
            vMail.To.Add(pRept_ID);

            //메일제목.
            vMail.Subject = isMessageAdapter1.ReturnText("KHRM_10035", string.Format("&&PAY_YYYYMM:={0} &&WAGE_TYPE:={1}", W_PAY_YYYYMM.EditValue, W_WAGE_TYPE_NAME.EditValue));
            vMail.SubjectEncoding = System.Text.Encoding.UTF8;

            //메일 내용.
            string vMail_Body = isMessageAdapter1.ReturnText("SKEAPP_10219");
            vMail_Body = string.Format("{0}\r\r{1}", vMail_Body, isMessageAdapter1.ReturnText("KHRM_10035", string.Format("&&PAY_YYYYMM:={0} &&WAGE_TYPE:={1}", W_PAY_YYYYMM.EditValue, W_WAGE_TYPE_NAME.EditValue)));
            vMail_Body = string.Format("{0}\r\r{1}", vMail_Body, isMessageAdapter1.ReturnText("SKEAPP_10220"));
            vMail.Body = vMail_Body;
            vMail.BodyEncoding = System.Text.Encoding.UTF8;

            //첨부파일 첨부.
            if (pAttachment_Path != string.Empty)
            {
                System.Net.Mail.Attachment vAttachment;
                vAttachment = new System.Net.Mail.Attachment(pAttachment_Path);
                vAttachment.NameEncoding = System.Text.Encoding.UTF8;
                vMail.Attachments.Add(vAttachment);     //첨부파일 붙이기.
            }

            //mail svr 설정.
            try
            {
                SmtpClient vSmtp = new SmtpClient(iConv.ISNull(O_SMTP_SVR.EditValue), iConv.ISNumtoZero(O_SMTP_PORT.EditValue, 25));
                vSmtp.UseDefaultCredentials = false;    //시스템에 설정된 인증 정보를 사용하지 않는다.
                vSmtp.EnableSsl = false;                 //SSL을 사용한다. 
                vSmtp.DeliveryMethod = SmtpDeliveryMethod.Network;  //이걸 사용하지 않으면 NAVER에 인증을 받지 못한다.
                vSmtp.Credentials = new NetworkCredential(iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue).Replace(string.Format("@{0}", O_SMTP_SVR.EditValue), ""), iConv.ISNull(V_EMAIL_ACCOUNT_PWD.EditValue));
                vSmtp.Send(vMail);
            }
            catch (Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

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
                    XLPrinting_Main("PRINT"); // 출력 함수 호출

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                    // 인쇄 완료 메시지 출력
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("FILE"); // 출력 함수 호출
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0528_Load(object sender, EventArgs e)
        {
            SUB_STATUS(false, "PRINT");
            CB_STAMP.BringToFront();

            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            W_CORP_NAME.BringToFront();

            W_PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);

            // 그리드 부분 업데이트 처리 위함.
            IDA_MONTH_PAYMENT.FillSchema();
            IDA_MONTH_PAYMENT_EMAIL.FillSchema();

            //메일서버 정보.
            IDC_GET_MAIL_SMTP_SVR.ExecuteNonQuery();

            //E MAIL전송자 정보
            IDC_GET_EMAIL_SENDER.ExecuteNonQuery();
            V_EMAIL_ACCOUNT.EditValue = IDC_GET_EMAIL_SENDER.GetCommandParamValue("O_EMAIL_ACCOUNT");
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
            IGR_MONTH_PAYMENT.LastConfirmChanges();
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
            IGR_MONTH_PAYMENT.LastConfirmChanges();
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            isAppInterfaceAdv1.OnAppMessage("Select Row [0]");
            System.Windows.Forms.Application.DoEvents();
        }
        
        private void BTN_SELECT_Y_2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < IGR_MONTH_PAYMENT_EMAIL.RowCount; i++)
            {
                IGR_MONTH_PAYMENT_EMAIL.SetCellValue(i, IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("SELECT_CHECK_YN"), "Y");
            }
            IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
            IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

            string vMessageText = string.Format("Select Row [{0}]", IGR_MONTH_PAYMENT.RowCount);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();
        }

        private void BTN_SELECT_N_2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < IGR_MONTH_PAYMENT_EMAIL.RowCount; i++)
            {
                IGR_MONTH_PAYMENT_EMAIL.SetCellValue(i, IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("SELECT_CHECK_YN"), "N");
            }
            IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
            IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

            isAppInterfaceAdv1.OnAppMessage("Select Row [0]");
            System.Windows.Forms.Application.DoEvents();
        }

        private void BTN_SEND_EMAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            System.DateTime vStartTime = DateTime.Now;

            string vMessageText = string.Empty;

            string vBoxCheck = string.Empty;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            string vSUBJECT = string.Empty;
            string vBODY = string.Empty;
            string vNOTICE = string.Empty;
            string vBOTTOM = string.Empty;

            int vCountCheck = 0;
            object vObject = null;
            int vCountRow = IGR_MONTH_PAYMENT_EMAIL.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }
            
            IGR_MONTH_PAYMENT.LastConfirmChanges();
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            IGR_MONTH_PAYMENT_EMAIL.LastConfirmChanges();
            IDA_MONTH_PAYMENT_EMAIL.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT_EMAIL.Refillable = true;

            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;

            int vIndexWAGE_TYPE = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PAY_TYPE");

            int vIndexPRINT_TYPE = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PRINT_TYPE");
            int vIndexPAY_YYYYMM = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PAY_YYYYMM");
            int vIndexPERSON_ID = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PERSON_ID");
            int vIndexNAME = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("NAME");
            int vIndexPERSON_NUM = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PERSON_NUM");
            int vIDX_EMAIL_ID = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("EMAIL");
            int vIDX_PASSWORD = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("PASSWORD");

            int vIndexCheckBox = IGR_MONTH_PAYMENT_EMAIL.GetColumnToIndex("SELECT_CHECK_YN");
            string vCheckedString = IGR_MONTH_PAYMENT_EMAIL.GridAdvExColElement[vIndexCheckBox].CheckedString;

            if (iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue) == string.Empty)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(string.Format("Sender : {0}", isMessageAdapter1.ReturnText("FCM_10256")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                V_EMAIL_ACCOUNT_ID.Focus();
                return;
            }
            //if (iConv.ISNull(V_EMAIL_ACCOUNT_PWD.EditValue) == string.Empty)
            //{
            //    Application.UseWaitCursor = false;
            //    System.Windows.Forms.Cursor.Current = Cursors.Default;
            //    Application.DoEvents();

            //    MessageBoxAdv.Show(string.Format("Sender : {0}", isMessageAdapter1.ReturnText("EAPP_10143")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    V_EMAIL_ACCOUNT_PWD.Focus();
            //    return;
            //}

            //저장폴더//
            string vSavePath = System.Environment.CurrentDirectory + "\\Pdf";
            try
            {
                if (System.IO.Directory.Exists(vSavePath))
                    System.IO.Directory.Delete(vSavePath, true);
                
                System.IO.Directory.CreateDirectory(vSavePath);
            }
            catch
            {
                //
            }

            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("E-Mail Send Data Save Start");
            System.Windows.Forms.Application.DoEvents(); 
            
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            //try
            //{
            //    //-------------------------------------------------------------------------------------
            //    xlPrinting.OpenFileNameExcel = "HRMF0528_002.xlsx"; 
            //}
            //catch (System.Exception ex)
            //{
            //    vMessageText = ex.Message;
            //    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            //    System.Windows.Forms.Application.DoEvents();
            //}

            //메일 본문.
            IDA_EMAIL_DOC.Fill();
            foreach(DataRow vROW in IDA_EMAIL_DOC.CurrentRows)
            {
                if(iConv.ISNull(vROW["EMAIL_TYPE"]) == "HEADER")
                {
                    if (vSUBJECT == string.Empty)
                    {
                        vSUBJECT = String.Format("{0} {1}", W_PAY_YYYYMM.EditValue, vROW["EMAIL_DOC"]);
                    }
                    else
                    {
                        vSUBJECT = String.Format("{0}\r\n{1}", vSUBJECT, vROW["EMAIL_DOC"]);
                    }
                }
                else if(iConv.ISNull(vROW["EMAIL_TYPE"]) == "BODY")
                {
                    if (vBODY == string.Empty)
                    {
                        vBODY = String.Format("{0}", vROW["EMAIL_DOC"]);
                    }
                    else
                    {
                        vBODY = String.Format("{0}\r\n{1}", vBODY, vROW["EMAIL_DOC"]);
                    }
                }
                else if(iConv.ISNull(vROW["EMAIL_TYPE"]) == "NOTICE")
                {
                    if (vNOTICE == string.Empty)
                    {
                        vNOTICE = String.Format("{0}", vROW["EMAIL_DOC"]);
                    }
                    else
                    {
                        vNOTICE = String.Format("{0}\r\n{1}", vNOTICE, vROW["EMAIL_DOC"]);
                    }
                }
                else if (iConv.ISNull(vROW["EMAIL_TYPE"]) == "BOTTOM")
                {
                    if (vBOTTOM == string.Empty)
                    {
                        vBOTTOM = String.Format("{0}", vROW["EMAIL_DOC"]);
                    }
                    else
                    {
                        vBOTTOM = String.Format("{0}\r\n{1}", vBOTTOM, vROW["EMAIL_DOC"]);
                    }
                }
            }
             
            vCountCheck = 0;
            for (int vRow = 0; vRow < IGR_MONTH_PAYMENT_EMAIL.RowCount; vRow++)
            {                
                vObject = IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexCheckBox);
                vBoxCheck = ConvertString(vObject);
                if (vBoxCheck == "Y")
                {
                    IGR_MONTH_PAYMENT_EMAIL.CurrentCellMoveTo(vRow, vIndexCheckBox);
                    IGR_MONTH_PAYMENT_EMAIL.CurrentCellActivate(vRow, vIndexCheckBox);
                    IGR_MONTH_PAYMENT_EMAIL.Focus();

                    vMessageText = string.Format("E-Mail Sending...{0}({1})", IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexNAME),
                                                                        IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexPERSON_NUM));
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    //판넬 view.
                    SUB_STATUS(true, "PRINT");
                    V_PRINT.CheckedState = ISUtil.Enum.CheckedState.Checked;

                    //체크한 대상중에 인쇄대상건만 인쇄//
                    vCountCheck++;

                    //파일명.
                    string vPay_YYYYMM = iConv.ISNull(IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexPAY_YYYYMM));
                    string vPerson_Num = iConv.ISNull(IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexPERSON_NUM));
                    string vPassword = iConv.ISNull(IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIDX_PASSWORD));
                    string vSaveFileName = String.Format("{0}\\{1}_{2}.pdf", vSavePath, vPay_YYYYMM, vPerson_Num);
                    try
                    {
                        if (System.IO.File.Exists(vSaveFileName))
                            System.IO.File.Delete(vSaveFileName); 
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage("XLLine Deduction" + Ex.Message); 
                        System.Windows.Forms.Application.DoEvents(); 
                    }
                     
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(string.Format("{0} => Printing", vMessageText));

                    int vFileCnt = xlPrinting.WriteMain_EMAIL(vSaveFileName
                                                            , vPassword
                                                            , vRow
                                                            , IGR_MONTH_PAYMENT_EMAIL
                                                            , IDA_MONTH_ALLOWANCE_E
                                                            , IDA_MONTH_DEDUCTION_E
                                                            , IDA_MONTH_DUTY_E
                                                            , IDA_MONTH_OT_E
                                                            , CB_STAMP.CheckedString);
                    if (vFileCnt < 1)
                    {
                        SUB_STATUS(false, "PRINT"); 
                        System.Windows.Forms.Application.DoEvents(); 
                        return;
                    } 
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(string.Format("{0} => Printed", vMessageText));

                    V_SAVE_FILE.CheckedState = ISUtil.Enum.CheckedState.Checked;

                    IDC_SET_EMAIL_COUNT.SetCommandParamValue("P_PAY_YYYYMM", IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexPAY_YYYYMM));
                    IDC_SET_EMAIL_COUNT.SetCommandParamValue("P_WAGE_TYPE", IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexWAGE_TYPE));
                    IDC_SET_EMAIL_COUNT.SetCommandParamValue("P_PERSON_ID", IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIndexPERSON_ID));
                    IDC_SET_EMAIL_COUNT.ExecuteNonQuery();

                    //이메일 발송.
                    MailMessage vMail = new MailMessage();

                    //보내는 사람 메일주소.
                    vMail.From = new MailAddress(iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue));
                     
                    try
                    {
                        string Email_ID = iConv.ISNull(IGR_MONTH_PAYMENT_EMAIL.GetCellValue(vRow, vIDX_EMAIL_ID));

                        //받는 사람 메일주소.여러 사람에게 보낼경우 계속 추가하면 됨.
                        vMail.To.Clear();
                        vMail.To.Add(Email_ID); 

                        //메일제목.
                        vMail.Subject = vSUBJECT;
                        vMail.SubjectEncoding = System.Text.Encoding.UTF8;

                        //메일 내용.
                        vMail.Body = string.Format("{0}\r\n\r\n{1}\r\n\r\n{2}", vBODY, vNOTICE, vBOTTOM);
                        vMail.BodyEncoding = System.Text.Encoding.UTF8;

                        //첨부파일 첨부.
                        if (vSaveFileName != string.Empty)
                        {
                            System.Net.Mail.Attachment vAttachment;
                            vAttachment = new System.Net.Mail.Attachment(vSaveFileName);
                            vAttachment.NameEncoding = System.Text.Encoding.UTF8;
                            vMail.Attachments.Clear();
                            try
                            {
                                vMail.Attachments.Add(vAttachment);     //첨부파일 붙이기.
                            }
                            catch(Exception Ex)
                            { 
                                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Email Attachment Add Error :: " + Ex.Message);
                            }
                        }
                        else
                        {
                            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Email Attachment Empty");
                        }

                        //mail svr 설정. 
                        SmtpClient vSmtp = new SmtpClient(iConv.ISNull(O_SMTP_SVR.EditValue), iConv.ISNumtoZero(O_SMTP_PORT.EditValue, 25));
                        if (O_USER_AUTH_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
                        {
                            vSmtp.UseDefaultCredentials = false;    //시스템에 설정된 인증 정보를 사용하지 않는다.
                        }
                        else
                        {
                            vSmtp.UseDefaultCredentials = true;    //시스템에 설정된 인증 정보를 사용하지 않는다.
                        }

                        if (O_SSL_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
                        {
                            vSmtp.EnableSsl = true;                //SSL을 사용한다. 
                        }
                        else
                        {
                            vSmtp.EnableSsl = false;                //SSL을 사용한다. 
                        }
                        vSmtp.DeliveryMethod = SmtpDeliveryMethod.Network;  //이걸 사용하지 않으면 NAVER에 인증을 받지 못한다.
                        if (O_USER_AUTH_FLAG.CheckedState == ISUtil.Enum.CheckedState.Checked)
                        {
                            //사용자 인증시 사용할ID
                            string Sender_ID = iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue).Substring(0, iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue).LastIndexOf(@"@"));
                            vSmtp.Credentials = new NetworkCredential(Sender_ID, iConv.ISNull(V_EMAIL_ACCOUNT_PWD.EditValue));
                        }
                         
                        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(string.Format("{0} => Email Sending", vMessageText));
                        
                        V_EMAIL_SEND.CheckedState = ISUtil.Enum.CheckedState.Checked;
                        vSmtp.Send(vMail); 
                        
                        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(string.Format("{0} => Email Sended", vMessageText));
                    }
                    catch (Exception Ex)
                    {
                        SUB_STATUS(false, "PRINT");
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        System.Windows.Forms.Application.DoEvents();
                        isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Email Send Error :: " + Ex.Message);

                        xlPrinting.Dispose();
                        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    } 
                    vMail.Dispose();
                    SUB_STATUS(false, "PRINT");
                    IGR_MONTH_PAYMENT_EMAIL.SetCellValue(vRow, vIndexCheckBox, "N"); 
                }
            }
            //-------------------------------------------------------------------------------------
            try
            {
                xlPrinting.Dispose();
            }
            catch(Exception Ex)
            {
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent("Print File Closing Error :: " + Ex.Message);
            }
            //-------------------------------------------------------------------------------------
             
            vMessageText = string.Format("E-Mail Sending...Compeleted");
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        private void V_SUB_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SUB_STATUS(false, "PRINT");
        }

        private void BTN_EMAIL_TEXT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_GET_EMAIL_PERSON_DESC.SetCommandParamValue("W_EMAIL_TYPE", "SALARY");
            IDC_GET_EMAIL_PERSON_DESC.ExecuteNonQuery();

            S_EMAIL_ACCOUNT.EditValue = V_EMAIL_ACCOUNT.EditValue;
            S_EMAIL_ACCOUNT_ID.EditValue = V_EMAIL_ACCOUNT_ID.EditValue;
            S_EMAIL_ACCOUNT_PWD.EditValue = V_EMAIL_ACCOUNT_PWD.EditValue;

            SUB_STATUS(true, "EMAIL_TEXT");
        }

        private void V_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_SAVE_EMAIL_PERSON_DESC.SetCommandParamValue("W_EMAIL_TYPE", "SALARY"); 
            IDC_SAVE_EMAIL_PERSON_DESC.ExecuteNonQuery();
            string vStatus = iConv.ISNull(IDC_SAVE_EMAIL_PERSON_DESC.GetCommandParamValue("O_STATUS"));
            string vMessage = iConv.ISNull(IDC_SAVE_EMAIL_PERSON_DESC.GetCommandParamValue("O_MESSAGE"));
            if(vStatus == "F")
            {
                if(vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            if(iConv.ISNull(V_EMAIL_ACCOUNT.EditValue) != iConv.ISNull(S_EMAIL_ACCOUNT.EditValue))
            {
                V_EMAIL_ACCOUNT.EditValue = S_EMAIL_ACCOUNT.EditValue;
            }
            if (iConv.ISNull(V_EMAIL_ACCOUNT_ID.EditValue) != iConv.ISNull(S_EMAIL_ACCOUNT_ID.EditValue))
            {
                V_EMAIL_ACCOUNT_ID.EditValue = S_EMAIL_ACCOUNT_ID.EditValue;
            }
            if (iConv.ISNull(V_EMAIL_ACCOUNT_PWD.EditValue) != iConv.ISNull(S_EMAIL_ACCOUNT_PWD.EditValue))
            {
                V_EMAIL_ACCOUNT_PWD.EditValue = S_EMAIL_ACCOUNT_PWD.EditValue;
            }
            SUB_STATUS(false, "EMAIL_TEXT");
        }

        private void V_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            S_EMAIL_ACCOUNT.EditValue = string.Empty;
            S_EMAIL_ACCOUNT_ID.EditValue = string.Empty;
            S_EMAIL_ACCOUNT_PWD.EditValue = string.Empty;

            SUB_STATUS(false, "EMAIL_TEXT");
        }

        #endregion

        #region ----- Lookup Event ----- 

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
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