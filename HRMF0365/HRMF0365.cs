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

namespace HRMF0365
{
    public partial class HRMF0365 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.

        #endregion;

        #region ----- Constructor -----

        public HRMF0365(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)   //파견직관리
            {
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }

        #endregion;

        #region ----- Corp Type -----

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

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
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
            G_CORP_GROUP.BringToFront();
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iConv.ISNull(G_CORP_TYPE.EditValue, "1") == "1")
            {
                G_CORP_GROUP.Visible = false; //.Show();
                V_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G_CORP_TYPE.EditValue = V_RB_OWNER.RadioCheckedString;
            }
            else
            {
                G_CORP_GROUP.Visible = true; //.Show();
                if (iConv.ISNull(G_CORP_TYPE.EditValue) == "ALL")
                {
                    V_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ALL.RadioCheckedString;
                }
                else
                {
                    V_RB_ETC.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ETC.RadioCheckedString;
                }
            }
        }
         
        private void SEARCH_DB()
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_WORK_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_WORK_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_FR.Focus();
                return;
            }
              
            IGR_OT_APPROVE.LastConfirmChanges();
            IDA_OT_APPROVE.OraSelectData.AcceptChanges();
            IDA_OT_APPROVE.Refillable = true;

            IDA_OT_APPROVE.SetSelectParamValue("W_SOB_ID", -1);
            IDA_OT_APPROVE.Fill();

            V_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            ////주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화
            ////평일에는 근무후 시작일, 시작시를 수정할 일이 없어 수정 불가능하도록 설정
            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0;  //수정 불가능, 근무후 일자
            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능, 근무후 시간

            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;

            IDA_OT_APPROVE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
            IDA_OT_APPROVE.Fill(); 
            IGR_OT_APPROVE.Focus(); 
        }

        private void SEARCH_DB_Calendar(object pPerson_ID, object pWork_Date)
        {
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_FR", iDate.ISDate_Add(pWork_Date, -3));
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_TO", pWork_Date);
            IDA_WORK_CALENDAR_S.Fill();
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
            
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_OT_APPROVE.GetCellValue("OT_ID"));
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            int mIDX_SELECT_YN = IGR_OT_APPROVE.GetColumnToIndex("SELECT_FLAG");
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                BTN_OK.Enabled = false;
                BTN_CANCEL.Enabled = false;
                BTN_RETURN.Enabled = false;

                IGR_OT_APPROVE.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
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
                IGR_OT_APPROVE.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
        }

        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (IGR_OT_APPROVE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_OT_APPROVE.LastConfirmChanges();
            IDA_OT_APPROVE.OraSelectData.AcceptChanges();
            IDA_OT_APPROVE.Refillable = true;

            int vIDX_SELECT_YN = IGR_OT_APPROVE.GetColumnToIndex("SELECT_FLAG");
            int vIDX_OT_ID = IGR_OT_APPROVE.GetColumnToIndex("OT_ID");
            int vIDX_APPROVE_STATUS = IGR_OT_APPROVE.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_OT_APPROVE.RowCount; i++)
            {
                if (iConv.ISNull(IGR_OT_APPROVE.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    IDC_UPDATE_APPROVE.SetCommandParamValue("W_OT_ID", IGR_OT_APPROVE.GetCellValue(i, vIDX_OT_ID));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_STATUS", IGR_OT_APPROVE.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", IGR_OT_APPROVE.GetCellValue(i, vIDX_SELECT_YN)); 
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

            SEARCH_DB(); 
        }

        private bool Set_Update_Return(DateTime pSys_Date)
        {
            if (IGR_OT_APPROVE.RowCount < 1)
            {
                return false;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_OT_APPROVE.LastConfirmChanges();
            IDA_OT_APPROVE.OraSelectData.AcceptChanges();
            IDA_OT_APPROVE.Refillable = true;

            int vIDX_SELECT_YN = IGR_OT_APPROVE.GetColumnToIndex("SELECT_FLAG");
            int vIDX_OT_ID = IGR_OT_APPROVE.GetColumnToIndex("OT_ID");
            int vIDX_WORK_DATE = IGR_OT_APPROVE.GetColumnToIndex("WORK_DATE");
            int vIDX_PERSON_ID = IGR_OT_APPROVE.GetColumnToIndex("PERSON_ID");
            int vIDX_APPROVE_STATUS = IGR_OT_APPROVE.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_OT_APPROVE.RowCount; i++)
            {
                if (iConv.ISNull(IGR_OT_APPROVE.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_OT_ID", IGR_OT_APPROVE.GetCellValue(i, vIDX_OT_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_CHECK_YN", IGR_OT_APPROVE.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_WORK_DATE", IGR_OT_APPROVE.GetCellValue(i, vIDX_WORK_DATE));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_PERSON_ID", IGR_OT_APPROVE.GetCellValue(i, vIDX_PERSON_ID)); 
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_APPROVE_STATUS", IGR_OT_APPROVE.GetCellValue(i, vIDX_APPROVE_STATUS));
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
            IDC_GET_DATE.ExecuteNonQuery();
            object vLOCAL_DATE = iDate.ISGetDate(IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE")).ToShortDateString();

            // EMAIL 발송.
            IDC_EMAIL_SEND.SetCommandParamValue("P_GUBUN", V_EMAIL_STATUS.EditValue);
            IDC_EMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "OT");
            IDC_EMAIL_SEND.SetCommandParamValue("P_WORK_DATE", vLOCAL_DATE);
            IDC_EMAIL_SEND.SetCommandParamValue("P_REQ_DATE", vLOCAL_DATE);
            IDC_EMAIL_SEND.ExecuteNonQuery();
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
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "HR_OT");
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
            string vSTATUS = "F";
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
            string vSTATUS = "F";
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
        }

        #endregion;
          
        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {                    
                    if (IDA_OT_APPROVE.IsFocused)
                    {
                        IDA_OT_APPROVE.AddOver();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_OT_APPROVE.IsFocused)
                    {
                        IDA_OT_APPROVE.AddUnder();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    System.Windows.Forms.SendKeys.Send("{TAB}");
                    try
                    {
                        IDA_OT_APPROVE.Update(); 
                      
                    }
                    catch(Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OT_APPROVE.IsFocused)
                    {
                        IDA_OT_APPROVE.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OT_APPROVE.IsFocused)
                    {
                        IDA_OT_APPROVE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print) //인쇄버튼
                {
                    XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export) //엑셀파일 버튼
                {
                    XLPrinting1("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0365_Load(object sender, EventArgs e)
        {
            W_CORP_NAME.BringToFront();
            
            W_WORK_DATE_FR.EditValue = iDate.ISDate_Add(DateTime.Today, -31);
            W_WORK_DATE_TO.EditValue = iDate.ISGetDate();
            BTN_OK.BringToFront();
            RB_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_APPROVE_STATUS.EditValue = RB_N.RadioCheckedString;
            V_EMAIL_STATUS.EditValue = "N";
            Set_BTN_STATE();

            DefaultCorporation();
            Set_FTP_Info();

            IDA_OT_APPROVE.FillSchema(); 
        }

        private void V_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            string mAPPROVE_STATE = iConv.ISNull(V_APPROVE_STATUS.EditValue);
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                return;
            }
            for (int r = 0; r < IGR_OT_APPROVE.RowCount; r++)
            {
                IGR_OT_APPROVE.SetCellValue(r, IGR_OT_APPROVE.GetColumnToIndex("SELECT_FLAG"), V_SELECT_YN.CheckBoxString);
            }
            IGR_OT_APPROVE.LastConfirmChanges();
            IDA_OT_APPROVE.OraSelectData.AcceptChanges();
            IDA_OT_APPROVE.Refillable = true;
        }

        private void irbALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();  // 버튼 상태 변경.
            SEARCH_DB(); 
        }

        #endregion;

        #region ----- Button Event -----

        #region ----- Event Remark -----

        //private void ibtSELECT_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        //{//대상산출
        //    int mRECORD_COUNT = 0;

        //    if (isOT_Line_Check() == false)
        //    {
        //        return;
        //    }

        //    idcOT_LINE_COUNT.ExecuteNonQuery();
        //    mRECORD_COUNT = Convert.ToInt32(idcOT_LINE_COUNT.GetCommandParamValue("O_RECORD_COUNT"));
        //    if (mRECORD_COUNT != Convert.ToInt32(0))
        //    {
        //        ////&&VALUE 는(은) 이미 존재합니다. &&TEXT 하시기 바랍니다.
        //        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10044", "&&VALUE:=Request Number's Data(신청번호에 대한 라인자료)&&TEXT:=Search(조회)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        //return;

        //        //[2011-07-25]
        //        idaOT_HEADER.Cancel();
        //        //기준일자에 대한 연장근무 신청이 이미 존재 합니다. 신청No로 조회해 수정 하십시오!
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10301"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    WORK_DATE.EditValue = STD_DATE_0.EditValue;

        //    idaOT_LINE.Cancel();

        //    idaINSERT_PERSON.Fill();

        //    igrOT_LINE.BeginUpdate();
        //    for (int i = 0; i < idaINSERT_PERSON.OraDataSet().Rows.Count; i++)
        //    {
        //        idaOT_LINE.AddUnder();
        //        for (int j = 0; j < igrOT_LINE.GridAdvExColElement.Count - 1; j++)
        //        {
        //            igrOT_LINE.SetCellValue(i, j + 1, idaINSERT_PERSON.OraDataSet().Rows[i][j]);
        //        }
        //    }
        //    igrOT_LINE.EndUpdate();
        //    igrOT_LINE.CurrentCellMoveTo(0, 0);
        //    igrOT_LINE.CurrentCellActivate(0, 0);
        //    igrOT_LINE.Focus();
        //}

        #endregion;

        private void BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                V_EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                V_EMAIL_STATUS.EditValue = "A1_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                V_EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                V_EMAIL_STATUS.EditValue = "N";
            }

            Set_Update_Approve("OK");
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EMAIL STATUS.
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                V_EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                V_EMAIL_STATUS.EditValue = "A1_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                V_EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "C".ToString())
            {
                V_EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                V_EMAIL_STATUS.EditValue = "N";
            }
            Set_Update_Approve("CANCEL");
        }

        private void BTN_RETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_WORK_DATE_FR.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_FR.Focus();
                return;
            }
            if (W_WORK_DATE_TO.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_TO.Focus();
                return;
            }
            if (Convert.ToDateTime(W_WORK_DATE_FR.EditValue) > Convert.ToDateTime(W_WORK_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_FR.Focus();
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

            Form vHRMF0365_RETURN = new HRMF0365_RETURN(isAppInterfaceAdv1.AppInterface
                                                        , W_CORP_ID.EditValue
                                                        , vLOCAL_DATE 
                                                        );
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0365_RETURN, isAppInterfaceAdv1.AppInterface);
            dlgResultValue = vHRMF0365_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            vHRMF0365_RETURN.Dispose();

            SEARCH_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void BTN_SELECT_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                IDA_OT_APPROVE.Update();
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                return;
            }

            object vOT_ID = IGR_OT_APPROVE.GetCellValue("OT_ID");

            //Document Revision Update.
            if (iConv.ISNull(vOT_ID) == string.Empty)
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

            object vDOCUMENT_REV_NUM = string.Format("{0}_{1:yyyyMMdd}", IGR_OT_APPROVE.GetCellValue("PERSON_NUM"), IGR_OT_APPROVE.GetCellValue("WORK_DATE"));

            if (UpLoadFile(vOT_ID, vDOCUMENT_REV_NUM) == true)
            {
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, vOT_ID);
            }
            SEARCH_DB();
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
            object vOT_ID = IGR_OT_APPROVE.GetCellValue("OT_ID");
            if (iConv.ISNull(vOT_ID) == string.Empty)
            {
                return;
            }
            //IDC_GET_DOC_LAST_REV_FLAG.SetCommandParamValue("W_DOC_REV_ID", vOT_ID);
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

        #endregion;

        #region ----- Adapter Event ------

        private void IDA_OT_APPROVE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                SEARCH_DB_Calendar(0, iDate.ISGetDate("1900-01-01"));
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, -1);
            }
            else
            {
                SEARCH_DB_Calendar(pBindingManager.DataRow["PERSON_ID"], pBindingManager.DataRow["WORK_DATE"]);
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_OT_APPROVE.GetCellValue("OT_ID"));
            }
        }

        #endregion

        #region ----- LookUP Event ----

        private void ILA_WORK_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
        }

        private void ILA_APPROVAL_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS"); 
        }

        private void ILA_PERSON_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_W.SetLookupParamValue("W_END_DATE", W_WORK_DATE_TO.EditValue);
        }
        
        private void ILA_DUTY_MANAGER_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_MANAGER.SetLookupParamValue("W_END_DATE", W_WORK_DATE_TO.EditValue);
            ILD_DUTY_MANAGER.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
            ILD_DUTY_MANAGER.SetLookupParamValue("W_CAP_CHECK_YN", "Y"); 
        }

        private void ilaPERSON_SelectedRowData(object pSender)
        {
            System.Windows.Forms.SendKeys.Send("{TAB}");
        }
          
        #endregion
         
        #region ----- XL Print 1 Method ----

        private void XLPrinting1(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //프린트 데이터 조회//
            IDA_PRINT_OT_REQ.Fill();
            int vCountRow = IDA_PRINT_OT_REQ.CurrentRows.Count;
             
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                vMessageText = string.Format("Printing File Open...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                string vREQ_PERSON_NAME = string.Format("신청자 : {0}", ""); //신청자
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0365_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                if (isOpen == true)
                {
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    vPageNumber = xlPrinting.XLWirteMain(IDA_PRINT_OT_REQ, vLOCAL_DATE, vREQ_PERSON_NAME);

                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.Save("OT_");
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