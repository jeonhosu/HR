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

namespace HRMF0305
{
    public partial class HRMF0305 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.

        #endregion;

        #region ----- Constructor -----

        public HRMF0305(Form pMainForm, ISAppInterface pAppInterface)
        {
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime( END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, 0);
            CB_SELECT.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            IDA_DUTY_PERIOD.Fill();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pStart_Date, Object pEnd_Date)
        {            
            idaWORK_CALENDAR.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            idaWORK_CALENDAR.SetSelectParamValue("W_WORK_DATE_FR", pStart_Date);
            idaWORK_CALENDAR.SetSelectParamValue("W_WORK_DATE_TO", pEnd_Date);

            if (pStart_Date != DBNull.Value && pEnd_Date != DBNull.Value)
            {
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_START_YEAR", iDate.ISYear(Convert.ToDateTime(pStart_Date)));
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_END_YEAR", iDate.ISYear(Convert.ToDateTime(pEnd_Date)));
            }
            idaWORK_CALENDAR.Fill();
            idaHOLIDAY_MANAGEMENT.Fill();
        }

        private bool isAdd_DB_Check()
        {// 데이터 추가시 검증.
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return false;
            }
            return true;
        }

        private void SEARCH_DB_ATTACHMENT(object pSOURCE_CATEGORY, object pSOURCE_ID)
        {
            //이미지 초기화;
            //ImageView(string.Empty);

            //첨부파일 리스트 조회 
            IDA_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_CATEGORY", pSOURCE_CATEGORY);
            IDA_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_ID", pSOURCE_ID);
            IDA_DOC_ATTACHMENT.Fill();
        }

        private void DELETE_DOC_ATTACHMENT()
        {
            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            if (iConv.ISNull(vDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DeleteFile(vDOC_ATTACHMENT_ID) == false)
            {
                return;
            }

            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_DUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID"));
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            int mIDX_SELECT_YN = IGR_DUTY_PERIOD.GetColumnToIndex("SELECT_FLAG");
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                BTN_OK.Enabled = false;
                BTN_CANCEL.Enabled = false;
                BTN_RETURN.Enabled = false;

                IGR_DUTY_PERIOD.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                if (mAPPROVE_STATE == "N")
                {
                    BTN_OK.Enabled = true;
                    BTN_CANCEL.Enabled = false;
                }
                else
                {
                    BTN_OK.Enabled = false;
                    BTN_CANCEL.Enabled = true;
                }
                BTN_RETURN.Enabled = true; 
                IGR_DUTY_PERIOD.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
        }

        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (IGR_DUTY_PERIOD.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_FLAG = IGR_DUTY_PERIOD.GetColumnToIndex("SELECT_FLAG");
            int vIDX_DUTY_PERIOD_ID = IGR_DUTY_PERIOD.GetColumnToIndex("DUTY_PERIOD_ID");
            int vIDX_APPROVE_STATUS = IGR_DUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DUTY_PERIOD.RowCount; i++)
            {
                if (iConv.ISNull(IGR_DUTY_PERIOD.GetCellValue(i, vIDX_SELECT_FLAG), "N") == "Y")
                {

                    IDC_UPDATE_APPROVE.SetCommandParamValue("W_DUTY_PERIOD_ID", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_DUTY_PERIOD_ID));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_STATUS", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_SELECT_FLAG));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_FLAG", pApproved_Flag);
                    IDC_UPDATE_APPROVE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            // eMail 전송.
            Send_Mail();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }


        private bool Set_Update_Return(DateTime pSys_Date)
        {
            if (IGR_DUTY_PERIOD.RowCount < 1)
            {
                return false;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_DUTY_PERIOD.LastConfirmChanges();
            IDA_DUTY_PERIOD.OraSelectData.AcceptChanges();
            IDA_DUTY_PERIOD.Refillable = true;
            
            int vIDX_SELECT_YN = IGR_DUTY_PERIOD.GetColumnToIndex("SELECT_FLAG");
            int vIDX_DUTY_PERIOD_ID = IGR_DUTY_PERIOD.GetColumnToIndex("DUTY_PERIOD_ID");
            int vIDX_START_DATE = IGR_DUTY_PERIOD.GetColumnToIndex("START_DATE");
            int vIDX_END_DATE = IGR_DUTY_PERIOD.GetColumnToIndex("END_DATE");
            int vIDX_PERSON_ID = IGR_DUTY_PERIOD.GetColumnToIndex("PERSON_ID");
            int vIDX_APPROVE_STATUS = IGR_DUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DUTY_PERIOD.RowCount; i++)
            {
                if (iConv.ISNull(IGR_DUTY_PERIOD.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_DUTY_PERIOD_ID", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_DUTY_PERIOD_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_CHECK_YN", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_START_DATE", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_START_DATE));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_END_DATE", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_END_DATE));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_PERSON_ID", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_PERSON_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_APPROVE_STATUS", IGR_DUTY_PERIOD.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_SYS_DATE", pSys_Date);
                    IDC_UPDATE_RETURN_TEMP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_RETURN_TEMP.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return false;
                    }
                }
            }
            return true;
        }

        private void Send_Mail()
        {
            IDC_GetDate.ExecuteNonQuery();
            object vLOCAL_DATE = iDate.ISGetDate(IDC_GetDate.GetCommandParamValue("X_LOCAL_DATE")).ToShortDateString();

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "DUTY");
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion;


        #region ----- FTP Infomation ----- 
        //ftp 접속정보 및 환경 정보 설정 
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            mFTP_Connect_Status = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "HR_DUTY");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                mFTP_Info = new isFTP_Info();

                mFTP_Info.Host = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mFTP_Info.Port = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mFTP_Info.UserID = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mFTP_Info.Password = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mFTP_Info.Passive_Flag = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));
                mFTP_Info.FTP_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mFTP_Info.Client_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mFTP_Info.Host == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            try
            {
                //FileTransfer Initialze
                mFileTransfer = new ISFileTransferAdv();
                mFileTransfer.Host = mFTP_Info.Host;
                mFileTransfer.Port = mFTP_Info.Port;
                mFileTransfer.UserId = mFTP_Info.UserID;
                mFileTransfer.Password = mFTP_Info.Password;
                if (mFTP_Info.Passive_Flag == "Y")
                {
                    mFileTransfer.UsePassive = true;
                }
                else
                {
                    mFileTransfer.UsePassive = false;
                }
                mDownload_Folder = string.Format("{0}\\{1}", mClient_Base_Path, mFTP_Info.Client_Folder);
            }
            catch (System.Exception Ex)
            {
                //ftp접속정보 오류 
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //Client Download Folder 없으면 생성 
            System.IO.DirectoryInfo vDownload_Folder = new System.IO.DirectoryInfo(mDownload_Folder);
            if (vDownload_Folder.Exists == false) //있으면 True, 없으면 False
            {
                vDownload_Folder.Create();
            }

            mFTP_Connect_Status = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
        //ftp에 file upload 처리 
        private bool UpLoadFile(object pDOC_REV_ID, object pDOCUMENT_REV_NUM)
        {
            bool isUpload = false;

            if (mFTP_Connect_Status == false)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iConv.ISNull(pDOCUMENT_REV_NUM) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "All File(*.*)|*.*|Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|PowerPoint File(*.ppt;*.pptx)|*.ppt;*.pptx|jpg file(*.jpg)|*.jpg|Pdf File(*.pdf)|*.pdf";
                openFileDialog1.DefaultExt = "*.*";
                openFileDialog1.FileName = "";
                openFileDialog1.Multiselect = false;

                //openFileDialog1.Title = "Select Open File";
                //openFileDialog1.Filter = "jpg file(*.jpg)|*.jpg|bmp file(*.bmp)|*.bmp";
                //openFileDialog1.DefaultExt = "jpg";
                //openFileDialog1.FileName = "";
                //openFileDialog1.Multiselect = false;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    //1. 사용자 선택 파일 
                    string vSelectFullPath = openFileDialog1.FileName;
                    string vSelectDirectoryPath = Path.GetDirectoryName(openFileDialog1.FileName);

                    string vFileName = Path.GetFileName(openFileDialog1.FileName);
                    string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();

                    //transaction 이용하기 위해 설정
                    IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = isDataTransaction1;
                    IDC_INSERT_DOC_ATTACHMENT.DataTransaction = isDataTransaction1;

                    //2. 첨부파일 DB 저장
                    isDataTransaction1.BeginTran();

                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_CATEGORY", V_DOC_CATEGORY.EditValue); //구분 
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_ID", pDOC_REV_ID);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_FTP_FILE_NAME", pDOCUMENT_REV_NUM);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                    IDC_INSERT_DOC_ATTACHMENT.ExecuteNonQuery();

                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));
                    object vDOC_ATTACHMENT_ID = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
                    object vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_FTP_FILE_NAME");

                    //O_DOC_ATTACHMENT_ID.EditValue = vDOC_ATTACHMENT_ID;
                    //O_FTP_FILE_NAME.EditValue = vFTP_FILE_NAME;

                    if (IDC_INSERT_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        isDataTransaction1.RollBack();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        //Transaction 해제.
                        IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                        IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                        return isUpload;
                    }

                    //3. 첨부파일 로그 저장 
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                    IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        isDataTransaction1.RollBack();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        //Transaction 해제.
                        IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                        IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                        return isUpload;
                    }

                    //4. 파일 업로드
                    try
                    {
                        int vArryCount = openFileDialog1.FileNames.Length;
                        for (int r = 0; r < vArryCount; r++)
                        {
                            mFileTransfer.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                            mFileTransfer.SourceFileName = vFileName;
                            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
                            mFileTransfer.TargetFileName = iConv.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransfer.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false;
                                isDataTransaction1.RollBack();
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                //Transaction 해제.
                                IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                                return isUpload;
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        isDataTransaction1.RollBack();
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                        return isUpload;
                    }

                    //5. 적용
                    isDataTransaction1.Commit();
                    //Transaction 해제.
                    IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                    IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                }
            }
            return isUpload;
        }

        #endregion;

        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(string pFileName)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            //string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();            
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", O_DOC_ATTACHMENT_ID.EditValue);
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            //IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            //vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            //if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            //{
            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();

            //    isDataTransaction1.RollBack();
            //    if (vMESSAGE != string.Empty)
            //    {
            //        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //    return IsDownload;
            //}            

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFileName);
            string vClientFileName = string.Format("{0}", pFileName);

            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.SourceFileName = pFileName;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시
                    string vClientFullPath = string.Format("{0}\\{1}", mDownload_Folder, vClientFileName);  //원본

                    System.IO.File.Delete(vClientFullPath);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, vClientFullPath);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
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
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
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

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        //ftp file download 처리 
        private bool DownLoadFile(string pSAVE_FileName, string pFTP_FILE_NAME)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            //string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_FILE_ENTRY_ID", pFILE_ENTRY_ID);
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            //IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            //vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            //if (vSTATUS == "F")
            //{
            //    isDataTransaction1.RollBack();
            //    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFTP_FILE_NAME);
            string vClientFileName = string.Format("{0}", pSAVE_FileName);

            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FILE_NAME;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시
                    string vClientFullPath = string.Format("{0}", vClientFileName);  //원본

                    System.IO.File.Delete(vClientFullPath);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, vClientFullPath);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
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
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
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

            //isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;

        #region ----- is View file Method -----

        private string isDownload(object pFileName)
        {
            string vFileName = iConv.ISNull(pFileName);

            if (vFileName != string.Empty)
            {
                if (DownLoadFile(vFileName) == true)
                {
                    return string.Format("{0}\\{1}", mDownload_Folder, vFileName);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        private string isDownload(string pSAVE_FileName, string pFTP_FILE_NAME)
        {
            if (pSAVE_FileName != string.Empty && pFTP_FILE_NAME != string.Empty)
            {
                if (DownLoadFile(pSAVE_FileName, pFTP_FILE_NAME) == true)
                {
                    return string.Format("{0}", pSAVE_FileName);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        #endregion;


        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(object pDOC_ATTACHMENT_ID)
        {
            bool IsDelete = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            object vDOC_ATTACHMENT_ID = pDOC_ATTACHMENT_ID;
            string vFTP_FileName = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            if (iConv.ISNull(vDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }
            if (iConv.ISNull(vFTP_FileName) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            //transaction 이용하기 위해 설정
            IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = isDataTransaction1;
            IDC_DELETE_DOC_ATTACHMENT.DataTransaction = isDataTransaction1;

            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            isDataTransaction1.BeginTran();
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "DELETE");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                isDataTransaction1.RollBack();
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }

            //2. 파일 삭제 
            IDC_DELETE_DOC_ATTACHMENT.SetCommandParamValue("W_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
            IDC_DELETE_DOC_ATTACHMENT.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false;
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }

            //3. 실제 삭제  
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;  //삭제는 소스에 설정해야 삭제됨.
            mFileTransfer.SourceFileName = vFTP_FileName;
            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.TargetFileName = vFTP_FileName;

            IsDelete = mFileTransfer.Delete();
            if (IsDelete == false)
            {
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }
            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            //Transaction 해제.
            IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
            IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
            return IsDelete;
        }

        #endregion;

        //-- Report 관련 코드
        #region ----- XL Export Methods ----

        private void ExportXL(ISGridAdvEx pGrid)
        {
            string vMessage = string.Empty;
            int vCountRows = pGrid.RowCount;

            if (vCountRows > 0)
            {
                saveFileDialog1.Title = "Excel_Save";
                saveFileDialog1.FileName = "Ex_00";
                saveFileDialog1.DefaultExt = "xls";
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
                saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    System.Windows.Forms.Application.DoEvents();

                    string vsSaveExcelFileName = saveFileDialog1.FileName;

                    XLExport mExport = new XLExport();
                    int vTerritory = GetTerritory(pGrid.TerritoryLanguage);
                    bool vbXLSaveOK = mExport.ExcelExport(pGrid, vTerritory, vsSaveExcelFileName, this.Text, this);
                    if (vbXLSaveOK == true)
                    {
                        vMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                        isAppInterfaceAdv1.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    else
                    {
                        vMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                        isAppInterfaceAdv1.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
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

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1()
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0305_001.xls";
                xlPrinting.XLFileOpen();

                int vTerritory = GetTerritory(IGR_DUTY_PERIOD.TerritoryLanguage);
                string vPeriodFrom = START_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);
                string vPeriodTo = END_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);

                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);

                int viCutStart = this.Text.LastIndexOf("]") + 1;
                string vCaption = this.Text.Substring(0, viCutStart);

                //xlPrinting.Req_Person_Name = "신 청 자 : " + igrDUTY_PERIOD.GetCellValue("REQUEST_PERSON_NAME");

                int vPageNumber = xlPrinting.XLWirte(IGR_DUTY_PERIOD, vTerritory, vPeriodFrom, vPeriodTo, vUserName, vCaption);

                xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                //xlPrinting.Printing(3, 4);


                xlPrinting.Save("Cashier_"); //저장 파일명

                //xlPrinting.PreView();

                xlPrinting.Dispose();
                //-------------------------------------------------------------------------

                vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
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
                    if(IDA_DUTY_PERIOD.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DUTY_PERIOD.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_DUTY_PERIOD.IsFocused)
                    {
                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DUTY_PERIOD.IsFocused)
                    {
                        IDA_DUTY_PERIOD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (idaDUTY_PERIOD.IsFocused)
                    //{
                    //    idaDUTY_PERIOD.Delete();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print) //인쇄버튼
                {
                    XLPrinting1();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export) //엑셀파일 버튼
                {
                    /*if (idaDUTY_PERIOD.IsFocused == true) // 어뎁터가 하나 이상일 경우 else if문으로 사용
                    {
                        ExportXL(idaDUTY_PERIOD);
                    }
                    */

                    ExportXL(IGR_DUTY_PERIOD);

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0305_Load(object sender, EventArgs e)
        {
            START_DATE_0.EditValue = DateTime.Today.AddDays(-31);
            END_DATE_0.EditValue = DateTime.Today.AddDays(7);

            // CORP SETTING
            DefaultCorporation();
            Set_FTP_Info();
            BTN_OK.BringToFront();
            RB_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_APPROVE_STATUS.EditValue = RB_N.RadioCheckedString;
            
            IDA_DUTY_PERIOD.FillSchema(); 
        }

        private void HRMF0305_Shown(object sender, EventArgs e)
        {
            Set_BTN_STATE();

            //LOOKUP SETTING
            ildAPPROVE_STATUS.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ildAPPROVE_STATUS.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
            ildSEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            ildSEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - SEARCH_TYPE
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            iSEARCH_TYPE_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            iSEARCH_TYPE_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize] 
            EMAIL_STATUS.EditValue = "N";
        }

        private void igrDUTY_PERIOD_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {// 시작일자 또는 종료일자 변경시 근무계획 조회.
            if (e.ColIndex == IGR_DUTY_PERIOD.GetColumnToIndex("NAME"))
            {
                isSearch_WorkCalendar(IGR_DUTY_PERIOD.GetCellValue("PERSON_ID"), IGR_DUTY_PERIOD.GetCellValue("START_DATE"), IGR_DUTY_PERIOD.GetCellValue("END_DATE"));
            }

            if (e.ColIndex == IGR_DUTY_PERIOD.GetColumnToIndex("START_DATE"))
            {
                isSearch_WorkCalendar(IGR_DUTY_PERIOD.GetCellValue("PERSON_ID"), e.NewValue, IGR_DUTY_PERIOD.GetCellValue("END_DATE"));                
            }
            if (e.ColIndex == IGR_DUTY_PERIOD.GetColumnToIndex("END_DATE"))
            {
                isSearch_WorkCalendar(IGR_DUTY_PERIOD.GetCellValue("PERSON_ID"), IGR_DUTY_PERIOD.GetCellValue("START_DATE"), e.NewValue);
            }
        }

        private void igrDUTY_PERIOD_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_DUTY_PERIOD.GetColumnToIndex("SELECT_FLAG");
            if (e.ColIndex == vIDX_SELECT_FLAG)
            {
                IGR_DUTY_PERIOD.LastConfirmChanges();
                IDA_DUTY_PERIOD.OraSelectData.AcceptChanges();
                IDA_DUTY_PERIOD.Refillable = true;
            }
        }

        private void btnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 승인
            // EMAIL STATUS.
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            Set_Update_Approve("OK");
        }

        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 취소
            // EMAIL STATUS.
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            Set_Update_Approve("CANCEL");
        }

        private void btnRETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime(END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            DialogResult dlgResultValue;

            //작업일자 
            IDC_GET_LOCAL_DATETIME_P.ExecuteNonQuery();
            DateTime vLOCAL_DATE = iDate.ISGetDate(IDC_GET_LOCAL_DATETIME_P.GetCommandParamValue("X_LOCAL_DATE"));

            //반려대상 선택.
            if (Set_Update_Return(vLOCAL_DATE) == false)
            {
                return;
            }

            Form vHRMF0305_RETURN = new HRMF0305_RETURN(isAppInterfaceAdv1.AppInterface
                                                        , CORP_ID_0.EditValue
                                                        , vLOCAL_DATE
                                                        );
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0305_RETURN, isAppInterfaceAdv1.AppInterface);
            dlgResultValue = vHRMF0305_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            vHRMF0305_RETURN.Dispose();

            Search_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }
        
        private void irbALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();  // 버튼 상태 변경.
            Search_DB();
        }

        private void CB_SELECT_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                return;
            }
            for (int r = 0; r < IGR_DUTY_PERIOD.RowCount; r++)
            {
                IGR_DUTY_PERIOD.SetCellValue(r, IGR_DUTY_PERIOD.GetColumnToIndex("SELECT_FLAG"), CB_SELECT.CheckBoxString);
            }
            IGR_DUTY_PERIOD.LastConfirmChanges();
            IDA_DUTY_PERIOD.OraSelectData.AcceptChanges();
            IDA_DUTY_PERIOD.Refillable = true;
        }

        private void BTN_SELECT_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                IDA_DUTY_PERIOD.Update();
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                return;
            }

            object vDUTY_PERIOD_ID = IGR_DUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID");
            object vPERSON_NUM = IGR_DUTY_PERIOD.GetCellValue("PERSON_NUM");
            object vSTART_DATE = IGR_DUTY_PERIOD.GetCellValue("START_DATE");
            object vEND_DATE = IGR_DUTY_PERIOD.GetCellValue("END_DATE"); 
            if (iConv.ISNull(vDUTY_PERIOD_ID) == string.Empty)
            {
                return;
            }
            if (iConv.ISNull(vPERSON_NUM) == string.Empty)
            {
                return;
            }
            if (iConv.ISNull(vSTART_DATE) == string.Empty)
            {
                return;
            }
            if (iConv.ISNull(vEND_DATE) == string.Empty)
            {
                return;
            }

            //IDC_GET_DOC_LAST_REV_FLAG.SetCommandParamValue("W_DOC_REV_ID", IGR_DOC_REVISION.GetCellValue("DOC_REV_ID"));
            //IDC_GET_DOC_LAST_REV_FLAG.ExecuteNonQuery();
            //string vLAST_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_REV_FLAG.GetCommandParamValue("O_LAST_REV_FLAG"));
            //if (vLAST_FLAG == "N")
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            object vDOCUMENT_REV_NUM = string.Format("{0}_{1:yyyyMMdd}_{2:yyyyMMdd}", vPERSON_NUM, vSTART_DATE, vEND_DATE);
            if (UpLoadFile(vDUTY_PERIOD_ID, vDOCUMENT_REV_NUM) == true)
            {
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, vDUTY_PERIOD_ID);
            }
            // SEARCH_DB();
        }

        private void BTN_FILE_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            ////업로드 가능자 아니면 최종버전만 다운로드 가능하도록 제어//
            //if (mMANAGER_FLAG != "Y")
            //{
            //    IDC_GET_DOC_LAST_REV_FLAG.ExecuteNonQuery();
            //    string vLAST_REV_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_REV_FLAG.GetCommandParamValue("O_LAST_REV_FLAG"));
            //    if (vLAST_REV_FLAG != "Y")
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10174"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}

            if (mFTP_Connect_Status == false)
            {
                MessageBoxAdv.Show("FTP IP is not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 저장될 Dialog 열기
            saveFileDialog1.Title = "Select Save Folder";
            saveFileDialog1.FileName = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = "C:\\";
            saveFileDialog1.Filter = "All file(*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string vSAVE_FILE_NAME = saveFileDialog1.FileName;
                string vFTP_FILE_NAME = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
                try
                {
                    isDownload(vSAVE_FILE_NAME, vFTP_FILE_NAME);
                }
                catch
                {
                    MessageBox.Show("Error : Could not read file from disk.");

                }

                System.Diagnostics.Process.Start(vSAVE_FILE_NAME);
            }
        }

        private void BTN_FILE_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vDUTY_PERIOD_ID = IGR_DUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID");
            if (iConv.ISNull(vDUTY_PERIOD_ID) == string.Empty)
            {
                return;
            }
            //IDC_GET_DOC_LAST_REV_FLAG.SetCommandParamValue("W_DOC_REV_ID", vDUTY_PERIOD_ID);
            //IDC_GET_DOC_LAST_REV_FLAG.ExecuteNonQuery();
            //string vLAST_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_REV_FLAG.GetCommandParamValue("O_LAST_REV_FLAG"));
            //if (vLAST_FLAG == "N")
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10168"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            DELETE_DOC_ATTACHMENT();
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDUTY_PERIOD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if(e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=사원 정보"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["START_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=시작일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["END_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=종료일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Convert.ToDateTime(e.Row["START_DATE"]) > Convert.ToDateTime(e.Row["END_DATE"]))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (String.IsNullOrEmpty(e.Row["DESCRIPTION"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=사유"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDUTY_PERIOD_PreDelete(ISPreDeleteEventArgs e)
        {
        }

        private void idaDUTY_PERIOD_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                isSearch_WorkCalendar(0, IGR_DUTY_PERIOD.GetCellValue("START_DATE"), IGR_DUTY_PERIOD.GetCellValue("END_DATE"));
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, 0);
                return;
            }
            isSearch_WorkCalendar(IGR_DUTY_PERIOD.GetCellValue("PERSON_ID"), IGR_DUTY_PERIOD.GetCellValue("START_DATE"), IGR_DUTY_PERIOD.GetCellValue("END_DATE"));
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_DUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID"));
        }

        #endregion

        #region ----- LookUp Event -----
        private void ilaAPPROVE_STATUS_0_SelectedRowData(object pSender)
        {
            IDA_DUTY_PERIOD.Fill();
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_CODE.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ildDUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaSTART_TIME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", IGR_DUTY_PERIOD.GetCellValue("START_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", "Y".ToString());
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", null);
        }

        private void ilaEND_TIME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", IGR_DUTY_PERIOD.GetCellValue("END_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", null);
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", "Y".ToString());
        }

        private void ilaSEARCH_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildSEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            ildSEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaJOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion
    }

    #region ----- FTP 정보 위한 사용자 Class -----

    public class isFTP_Info
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = "N";
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public isFTP_Info()
        {

        }

        public isFTP_Info(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag, string pFTP_Folder, string pClient_Folder)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mPassive_Flag = pPassive_Flag;
            mFTP_Folder = pFTP_Folder;
            mClient_Folder = pClient_Folder;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string Passive_Flag
        {
            get
            {
                return mPassive_Flag;
            }
            set
            {
                mPassive_Flag = value;
            }
        }

        public string FTP_Folder
        {
            get
            {
                return mFTP_Folder;
            }
            set
            {
                mFTP_Folder = value;
            }
        }

        public string Client_Folder
        {
            get
            {
                return mClient_Folder;
            }
            set
            {
                mClient_Folder = value;
            }
        }

        #endregion;
    }

    #endregion

}