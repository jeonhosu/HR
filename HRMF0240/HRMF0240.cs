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

namespace HRMF0240
{
    
    public partial class HRMF0240 : Office2007Form
    { 
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private string m_User_Print = "N"; 
        private object m_Req_Date = null;
        private string m_Print_Num = "";
        private object m_Print_Date = null;

        private string mREPORT_TYPE = string.Empty;
        private string mREPORT_FILENAME = string.Empty;

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.
        private string mClientFile = string.Empty;    // 현재 디렉토리및파일.

        private bool mIsGetInformationFTP = false;
        private string mHost = string.Empty;
        private string mPort = "21";
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = "N";

        private string mFILE_NAME = string.Empty;
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = "Image";

        private float fSIZE_W = 0;
        private float fSIZE_H = 0;
        private float fLOC_X = 0;
        private float fLOC_Y = 0;

        public HRMF0240(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;

            CORP_ID.EditValue = 25;
            m_User_Print = "N";
            if (m_User_Print.Equals("Y"))
            {
                PRINT_REQ_NUM.EditValue = "재직-22-009";
                m_Req_Date = "2022-05-24";
                m_Print_Num = "";
                m_Print_Date = "";
            }
            else
            {
                PRINT_REQ_NUM.EditValue = "";
                m_Req_Date = "";
                m_Print_Num = "재직-22-009";
                m_Print_Date = "2022-05-24";
            }
        }

        public HRMF0240(Form pMainForm, ISAppInterface pAppInterface, object pCorp_ID)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
            CORP_ID.EditValue = pCorp_ID; 
            m_User_Print = "N";
        }

        public HRMF0240(Form pMainForm, ISAppInterface pAppInterface, object pCorp_ID, string pUser_Print, string pPrint_Req_Num, object pReq_Date)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
            CORP_ID.EditValue = pCorp_ID;
            PRINT_REQ_NUM.EditValue = pPrint_Req_Num;

            m_User_Print = pUser_Print; 
            m_Req_Date = pReq_Date;
            m_Print_Num = "";
            m_Print_Date = "";
        }

        public HRMF0240(Form pMainForm, ISAppInterface pAppInterface, object pCorp_ID, string pPrint_NUM, object pPrint_Date)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
            CORP_ID.EditValue = pCorp_ID; 
            PRINT_REQ_NUM.EditValue = "";

            m_User_Print = "N"; 
            m_Req_Date = "";
            m_Print_Num = pPrint_NUM;
            m_Print_Date = pPrint_Date;
        }


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

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting_Main(string pPRINT_TYPE)
        {
            if (pPRINT_TYPE.Equals("TEST"))
            {
                mREPORT_TYPE = "TEST";
                mREPORT_FILENAME = "HRMF0240_001.xlsx"; 
            }
            else
            {
                IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0240");
                IDC_GET_REPORT_SET_P.ExecuteNonQuery();
                mREPORT_TYPE = iConv.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
                mREPORT_FILENAME = iConv.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_FILE_NAME")); 
            }
            XLPrinting(pPRINT_TYPE);

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
         
        private void XLPrinting(string pPRINT_TYPE)
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting(); 
            try
            {
                //-------------------------------------------------------------------------
                //-------------------------------------------------------------------------
                if (mREPORT_FILENAME != String.Empty)
                {
                    xlPrinting.OpenFileNameExcel = mREPORT_FILENAME;
                }
                else
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0240_003.xlsx";
                } 
                xlPrinting.XLFileOpen();

                int vPageCnt = 0;
                string vPeriodFrom = PRINT_DATE.DateTimeValue.ToString("yyyy-MM-dd", null); 
                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);
                string vREPRE_FLAG = PRINT_REPRE_FLAG.CheckBoxString;
                string vHISTORY_FLAG = PRINT_PERSON_HISTORY.CheckBoxString;
                string vSTAMP_FLAG = PRINT_STAMP.CheckBoxString;
                if (vHISTORY_FLAG.Equals("Y"))
                    IDA_HISTORY_DATA.Fill();

                int nPrintTotalCnt = iConv.ISNumtoZero(PRINT_COUNT.EditValue);
                if (pPRINT_TYPE.Equals("TEST"))
                {
                    if (PRINT_PREVIEW.CheckedState == ISUtil.Enum.CheckedState.Checked)
                        xlPrinting.PreView(1, 1);
                    else
                        xlPrinting.Printing(1, 1); //시작 페이지 번호, 종료 페이지 번호
                }
                else
                {
                    //V_LANG_CODE.EditValue
                    vPageCnt = xlPrinting.XLWirte(IDA_CERTIFICATE_INFO, IDA_HISTORY_DATA, nPrintTotalCnt
                                                , vPeriodFrom, vUserName 
                                                , pPRINT_TYPE, vREPRE_FLAG, vHISTORY_FLAG
                                                , vSTAMP_FLAG, mClientFile, fSIZE_W, fSIZE_H, fLOC_X, fLOC_Y);
                    if (pPRINT_TYPE.Equals("PDF"))
                    {
                        //기본 저장 경로 지정.            
                        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                        string vSaveFileName = iConv.ISNull(PRINT_NUM.EditValue); //기본 파일명.수정필요.

                        saveFileDialog1.Title = "Pdf Save";
                        saveFileDialog1.FileName = vSaveFileName;
                        saveFileDialog1.Filter = "pdf File(*.pdf)"; //"xlsx File(*.xlsx)|*.xlsx|CSV file(*.csv)|*.csv|Excel file(*.xls)|*.xls";
                        saveFileDialog1.DefaultExt = "pdf";
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
                        xlPrinting.Save(vSaveFileName);
                        vMessageText = string.Format(" Writing Starting...");

                        System.Windows.Forms.Application.UseWaitCursor = true;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    else
                    {
                        if (PRINT_PREVIEW.CheckedState == ISUtil.Enum.CheckedState.Checked)
                            xlPrinting.PreView(1, vPageCnt);
                        else
                            xlPrinting.Printing(1, vPageCnt); //시작 페이지 번호, 종료 페이지 번호
                    }
                }  
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------

                vMessageText = string.Format("Print End! [Page : {0}]", vPageCnt);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        #endregion;


        #region ----- Get Information FTP Methods -----

        private bool GetInfomationFTP()
        {
            bool isGet = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "COMM_DOC");
                IDC_FTP_INFO.ExecuteNonQuery();

                mHost = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mPort = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mUserID = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mPassword = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mPassive_Flag = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));

                mFTP_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mClient_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
                mClient_Folder = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Folder);

#if DEBUG
                {
                    mHost = "106.251.238.98";
                    mPort = "1502";
                    mUserID = "infoftp";
                    mPassword = "Infof12X";
                    mPassive_Flag = "Y";

                    mFTP_Folder = "/HETN_PROD_R2/FILE/COMM_DOC";
                    mClient_Folder = "Image";
                    mClient_Folder = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Folder);
                }
#endif

                if (mHost != string.Empty)
                {
                    mFileTransferAdv = new ISFileTransferAdv();
                    mFileTransferAdv.Host = mHost;
                    mFileTransferAdv.Port = mPort;
                    mFileTransferAdv.UserId = mUserID;
                    mFileTransferAdv.Password = mPassword;
                    mFileTransferAdv.KeepAlive = false;
                    if (mPassive_Flag == "Y")
                    {
                        mFileTransferAdv.UsePassive = true;
                    }
                    else
                    {
                        mFileTransferAdv.UsePassive = false;
                    }

                    isGet = true;
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
            return isGet;
        }


        private void GET_CO_STAMP()
        {
            IDC_GET_CORP_STAMP_P.SetCommandParamValue("W_ASSEMBLY_ID", "HRMF0240");
            IDC_GET_CORP_STAMP_P.ExecuteNonQuery();
            //mFTP_Folder= iConv.ISNull(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_FTP_PATH"));
            mFILE_NAME = iConv.ISNull(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_FILE_NAME"));

            fSIZE_W = ((float)iConv.ISDecimaltoZero(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_SIZE_W"), 0));
            fSIZE_H = ((float)iConv.ISDecimaltoZero(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_SIZE_H"), 0));
            fLOC_X = ((float)iConv.ISDecimaltoZero(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_LOC_X"), 0));
            fLOC_Y = ((float)iConv.ISDecimaltoZero(IDC_GET_CORP_STAMP_P.GetCommandParamValue("O_STAMP_LOC_Y"), 0));

#if DEBUG
            {
                mFILE_NAME = "C01.PNG"; 
            }
#endif 
            if (mFILE_NAME.Equals(""))
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
            if (DownLoadFile(mFILE_NAME) == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
        }

        #endregion;


        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(string pFILE_NAME)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;

            System.IO.DirectoryInfo vClientFolder = new System.IO.DirectoryInfo(mClient_Folder);
            if (vClientFolder.Exists == false) //있으면 True, 없으면 False
            {
                vClientFolder.Create();
            }

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFILE_NAME);
            try
            {
                System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                if (vDownFileInfo.Exists == true)
                {
                    try
                    {
                        System.IO.File.Delete(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                    }
                    catch
                    {

                        // ignore
                    }
                }
            }
            catch
            {
                //ignore                        
            }

            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------
            mFileTransferAdv.SourceDirectory = mFTP_Folder;
            mFileTransferAdv.SourceFileName = pFILE_NAME;
            mFileTransferAdv.TargetDirectory = mClient_Folder;
            mFileTransferAdv.TargetFileName = vTempFileName;

            IsDownload = mFileTransferAdv.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    mClientFile = string.Format("{0}\\{1}", mClient_Folder, pFILE_NAME);      //임시
                    System.IO.File.Delete(mClientFile);                 //기존 파일 삭제 

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mClient_Folder, vTempFileName);      //임시
                    System.IO.File.Move(vTempFullPath, mClientFile);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                            }
                            catch
                            {

                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(string.Format("{0}\\_{1}", mClient_Folder, vTempFileName));
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }
            if (IsDownload != true)
            {
                string vMessage = string.Format("{0} {1}", isMessageAdapter1.ReturnText("EAPP_10212"), isMessageAdapter1.ReturnText("QM_10102"));
                MessageBoxAdv.Show(string.Format("{0}\r\n{1}\r\n{2}", vMessage, mHost, mClientFile), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;



        #region ----- Form Event -----

        private void HRMF0240_Load(object sender, EventArgs e)
        {
            V_RB_PRINT.CheckedState = ISUtil.Enum.CheckedState.Checked;
            PRINT_TYPE.EditValue = V_RB_PRINT.RadioCheckedString;

            PRINT_DATE.Focus();
        }

        private void HRMF0240_Shown(object sender, EventArgs e)
        {
            //기본값 설정//
            IDC_GET_PRINT_CERTIFICATE.SetCommandParamValue("W_USER_PRINT", m_User_Print);
            IDC_GET_PRINT_CERTIFICATE.SetCommandParamValue("W_PRINT_REQ_NUM", PRINT_REQ_NUM.EditValue);
            IDC_GET_PRINT_CERTIFICATE.SetCommandParamValue("W_PRINT_NUM", m_Print_Num);
            IDC_GET_PRINT_CERTIFICATE.SetCommandParamValue("P_CORP_ID", CORP_ID.EditValue);
            IDC_GET_PRINT_CERTIFICATE.ExecuteNonQuery();

            if (m_User_Print.Equals("Y"))
            {
                PRINT_DATE.ReadOnly = true;
                PRINT_DATE.Refresh();

                CERT_TYPE_NAME.ReadOnly = true;
                CERT_TYPE_NAME.Refresh();

                NAME.ReadOnly = true;
                NAME.Refresh();

                TASK_DESC.ReadOnly = true;
                TASK_DESC.Refresh();

                SEND_ORG.ReadOnly = true;
                SEND_ORG.Refresh();

                REMARK.ReadOnly = true;
                REMARK.Refresh();

                PRINT_COUNT.ReadOnly = true;
                PRINT_COUNT.Refresh();

                V_RB_PRINT.Visible = false;
                V_RB_PDF.Visible = false;
            }

            mIsGetInformationFTP = GetInfomationFTP();
            GET_CO_STAMP();
        }

        private void V_RB_PRINT_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_PRINT.Checked == true)
            {
                PRINT_TYPE.EditValue = V_RB_PRINT.RadioCheckedString;
            }
        }

        private void V_RB_PDF_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_PDF.Checked == true)
            {
                PRINT_TYPE.EditValue = V_RB_PDF.RadioCheckedString;
            }
        }

        private void BTN_PRINT_TEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            XLPrinting_Main("TEST");
        }

        private void ibtPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 증명서 발급
            if (CERT_TYPE_ID.EditValue == null)
            {// 증명서 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10033"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CERT_TYPE_NAME.Focus();
                return;
            }

            if (PERSON_ID.EditValue == null)
            {// 사원 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CERT_TYPE_NAME.Focus();
                return;
            }

            if (string.IsNullOrEmpty(REMARK.EditValue.ToString()))
            {// 용도 입력
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10034"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CERT_TYPE_NAME.Focus();
                return;
            }
             
            // 인쇄 결과 저장.     
            IDC_CERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_CORP_ID", CORP_ID.EditValue);
            IDC_CERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDC_CERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDC_CERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
            IDC_CERTIFICATE_PRINT_INSERT.ExecuteNonQuery();
            PRINT_NUM.EditValue = IDC_CERTIFICATE_PRINT_INSERT.GetCommandParamValue("P_PRINT_NUM");
#if DEBUG
            PRINT_NUM.EditValue = "재직-22-017";
#endif

            // 인쇄발급 루틴 추가 //
            if (iConv.ISNull(PRINT_NUM.EditValue) == string.Empty)
            {// 인쇄번호 없음. 인쇄 실패.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Print_Certificate(iedPRINT_NUM.EditValue); // 증명서 인쇄 폼 안에 있는 그리드 관련 함수

            //인쇄하기//
            IDA_CERTIFICATE_INFO.Fill(); // 증명서 인쇄 폼 내에 그리드 부분에 삽입될 데이터 처리.
            if(IDA_CERTIFICATE_INFO.CurrentRows.Count < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                return;
            } 
            XLPrinting_Main(iConv.ISNull(PRINT_TYPE.EditValue));

            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(isMessageAdapter1.ReturnText("FCM_10035"));
            // 인쇄 완료 메시지 출력

            PRINT_NUM.EditValue = null;
            PRINT_DATE.EditValue = null;
            CERT_TYPE_ID.EditValue = null;
            CERT_TYPE_NAME.EditValue = null;

            PERSON_ID.EditValue = null;
            PERSON_NUM.EditValue = null;
            NAME.EditValue = null;
            DEPT_NAME.EditValue = null;
            POST_NAME.EditValue = null; 
            JOIN_DATE.EditValue = null;            
            RETIRE_DATE.EditValue = null;
            REMARK.EditValue = null;
            SEND_ORG.EditValue = null;
            PRINT_COUNT.EditValue = 1;
            DESCRIPTION.EditValue = null;

            this.Close();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        private void ilaCERT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CERT_TYPE.SetLookupParamValue("W_USER_PRINT_YN", m_User_Print); 
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (EMPLOYE_TYPE.EditValue.ToString() == "1".ToString())
            {
                ILD_PERSON.SetLookupParamValue("W_START_DATE", PRINT_DATE.EditValue);
                ILD_PERSON.SetLookupParamValue("W_END_DATE", PRINT_DATE.EditValue);
            }
            else
            {
                ILD_PERSON.SetLookupParamValue("W_START_DATE", DateTime.Parse("2001-01-01"));
                ILD_PERSON.SetLookupParamValue("W_END_DATE", DateTime.Today);
            }
            ILD_PERSON.SetLookupParamValue("W_CORP_ID", CORP_ID.EditValue);
        }
        #endregion

    }
}