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

namespace HRMF0304
{
    public partial class HRMF0304 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.

        private bool mSave_Flag = false;

        #endregion;

        #region ----- Constructor -----

        public HRMF0304(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void isSearch_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iSTART_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iSTART_DATE_0.Focus();
                return;
            }
            if (iEND_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iEND_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(iSTART_DATE_0.EditValue) > Convert.ToDateTime( iEND_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iSTART_DATE_0.Focus();
                return;
            }

            igrDUTY_PERIOD.LastConfirmChanges();
            idaDUTY_PERIOD.OraSelectData.AcceptChanges();
            idaDUTY_PERIOD.Refillable = true;

            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, 0);
            idaDUTY_PERIOD.Fill();
            igrDUTY_PERIOD.Focus();
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
            
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, igrDUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID"));
        }

        private void Init_Approve_Status()
        {
            int mIDX_COL = igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS");
            for (int R = 0; R < igrDUTY_PERIOD.RowCount; R++)
            {
                if (iConv.ISNull(igrDUTY_PERIOD.GetCellValue(R, mIDX_COL)) == "R".ToString() &&
                    idaDUTY_PERIOD.OraSelectData.Rows[R].RowState != DataRowState.Unchanged)
                {// 승인미요청 건에 대해서 승인 처리.
                    igrDUTY_PERIOD.SetCellValue(R, mIDX_COL, "N");
                }
            }
        }

        #endregion;

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
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                        return isUpload;
                    } 
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
             
            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리  
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
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            } 

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents(); 
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
                xlPrinting.OpenFileNameExcel = "HRMF0304_001.xls";
                xlPrinting.XLFileOpen();

                int vTerritory = GetTerritory(igrDUTY_PERIOD.TerritoryLanguage);
                string vPeriodFrom = iSTART_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);
                string vPeriodTo = iEND_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);

                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);

                int viCutStart = this.Text.LastIndexOf("]") + 1;
                string vCaption = this.Text.Substring(0, viCutStart);

                xlPrinting.Req_Person_Name = "신 청 자 : " + igrDUTY_PERIOD.GetCellValue("REQUEST_PERSON_NAME");

                int vPageNumber = xlPrinting.XLWirte(igrDUTY_PERIOD, vTerritory, vPeriodFrom, vPeriodTo, vUserName, vCaption);

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
                    isSearch_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if(idaDUTY_PERIOD.IsFocused)
                    {
                        if (isAdd_DB_Check() == false)
                        {
                            return;
                        }

                        idaDUTY_PERIOD.AddOver();

                        igrDUTY_PERIOD.SetCellValue("SELECT_YN", "Y");
                        igrDUTY_PERIOD.SetCellValue("START_DATE", DateTime.Today.Date);
                        igrDUTY_PERIOD.SetCellValue("END_DATE", DateTime.Today.Date);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaDUTY_PERIOD.IsFocused)
                    {
                        if (isAdd_DB_Check() == false)
                        {
                            return;
                        }
                        idaDUTY_PERIOD.AddUnder();

                        igrDUTY_PERIOD.SetCellValue("SELECT_YN", "Y");
                        igrDUTY_PERIOD.SetCellValue("START_DATE", DateTime.Today.Date);
                        igrDUTY_PERIOD.SetCellValue("END_DATE", DateTime.Today.Date);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    Init_Approve_Status();
                    idaDUTY_PERIOD.Update();                        
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    idaDUTY_PERIOD.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDUTY_PERIOD.IsFocused)
                    {
                        idaDUTY_PERIOD.Delete();
                    }
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

                    ExportXL(igrDUTY_PERIOD);

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0304_Load(object sender, EventArgs e)
        {            
            iSTART_DATE_0.EditValue = DateTime.Today.AddDays(-31);
            iEND_DATE_0.EditValue = DateTime.Today.AddDays(7);

            W_START_DATE.EditValue = DateTime.Today.Date;
            W_END_DATE.EditValue = DateTime.Today.Date;

            DefaultCorporation();
            Set_FTP_Info();

            // LOOKUP DEFAULT VALUE SETTING - DUTY_APPROVE_STATUS
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            iAPPROVE_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            iAPPROVE_STATUS_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            // LOOKUP DEFAULT VALUE SETTING - SEARCH_TYPE
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            iSEARCH_TYPE_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            iSEARCH_TYPE_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            idaDUTY_PERIOD.FillSchema();
        }

        private void igrDUTY_PERIOD_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {// 시작일자 또는 종료일자 변경시 근무계획 조회.
            if (e.ColIndex == igrDUTY_PERIOD.GetColumnToIndex("NAME"))
            {
                isSearch_WorkCalendar(igrDUTY_PERIOD.GetCellValue("PERSON_ID"), igrDUTY_PERIOD.GetCellValue("START_DATE"), igrDUTY_PERIOD.GetCellValue("END_DATE"));
            }

            if (e.ColIndex == igrDUTY_PERIOD.GetColumnToIndex("START_DATE"))
            {
                igrDUTY_PERIOD.SetCellValue("END_DATE", e.NewValue);
                isSearch_WorkCalendar(igrDUTY_PERIOD.GetCellValue("PERSON_ID"), e.NewValue, igrDUTY_PERIOD.GetCellValue("END_DATE"));                
            }
            if (e.ColIndex == igrDUTY_PERIOD.GetColumnToIndex("END_DATE"))
            {
                isSearch_WorkCalendar(igrDUTY_PERIOD.GetCellValue("PERSON_ID"), igrDUTY_PERIOD.GetCellValue("START_DATE"), e.NewValue);
            }
        }
        
        private void btnAPPR_REQUEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            mSave_Flag = true;
            Init_Approve_Status();
            try
            {
                idaDUTY_PERIOD.Update();
            }
            catch(Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (mSave_Flag == false)
            {
                return;
            } 

            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10240"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            string vAPPROVE_STATUS = null;
            object vDUTY_PERIOD_ID = string.Empty;
            int vIDX_DUTY_PERIOD_ID = igrDUTY_PERIOD.GetColumnToIndex("DUTY_PERIOD_ID");
            int mModifyCount = 0;
            int mRowCount = igrDUTY_PERIOD.RowCount;
            string vSELECT_YN = string.Empty; 
            for (int R = 0; R < mRowCount; R++)
            {
                vSELECT_YN = igrDUTY_PERIOD.GetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("SELECT_YN")).ToString();
                vDUTY_PERIOD_ID = igrDUTY_PERIOD.GetCellValue(R, vIDX_DUTY_PERIOD_ID);
                if ( vSELECT_YN == "Y")
                {
                    vAPPROVE_STATUS = iConv.ISNull(igrDUTY_PERIOD.GetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS")), "N");
                    if (vAPPROVE_STATUS == "N".ToString() || vAPPROVE_STATUS == "R".ToString())
                    {// 승인미요청 건에 대해서 승인 처리.
                        mModifyCount = mModifyCount + 1;

                        IDC_APPROVAL_REQUEST.SetCommandParamValue("W_DUTY_PERIOD_ID", vDUTY_PERIOD_ID);
                        IDC_APPROVAL_REQUEST.ExecuteNonQuery();
                        vSTATUS = iConv.ISNull(IDC_APPROVAL_REQUEST.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_APPROVAL_REQUEST.GetCommandParamValue("O_MESSAGE"));
                        if(vSTATUS == "F")
                        {
                            if(vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        object mValue;
                        mValue = IDC_APPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                        igrDUTY_PERIOD.SetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS"), mValue);
                        mValue = IDC_APPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                        igrDUTY_PERIOD.SetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                    }
                } 
            }

            // EMAIL 발송.
            if (mModifyCount > 0)
            {
                IDC_GetDate.ExecuteNonQuery();
                object vLOCAL_DATE = iDate.ISGetDate(IDC_GetDate.GetCommandParamValue("X_LOCAL_DATE")).ToShortDateString();

                idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
                idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "DUTY");
                idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
                idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", vLOCAL_DATE);
                idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", vLOCAL_DATE);
                idcEMAIL_SEND.ExecuteNonQuery();

                idaDUTY_PERIOD.OraSelectData.AcceptChanges();
                idaDUTY_PERIOD.Refillable = true;
            }

            igrDUTY_PERIOD.LastConfirmChanges();
            idaDUTY_PERIOD.OraSelectData.AcceptChanges();
            idaDUTY_PERIOD.Refillable = true;
            isSearch_DB();
        }

        private void btnAPPR_R_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //Init_Approve_Status();
            //idaDUTY_PERIOD.Update();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            string vAPPROVE_STATUS = null;
            int mModifyCount = 0;
            int mRowCount = igrDUTY_PERIOD.RowCount;
            string vSELECT_YN = string.Empty;
            for (int R = 0; R < mRowCount; R++)
            {
                vSELECT_YN = igrDUTY_PERIOD.GetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("SELECT_YN")).ToString();
                if (vSELECT_YN == "Y")
                {
                    vAPPROVE_STATUS = iConv.ISNull(igrDUTY_PERIOD.GetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS")));
                    if (vAPPROVE_STATUS == "A".ToString() /*|| vAPPROVE_STATUS == "R".ToString()*/)
                    {// 승인요청건에 대해서 승인미처리
                        mModifyCount = mModifyCount + 1;

                        IDC_DATA_UPDATE_REQUEST_CANCEL.SetCommandParamValue("W_DUTY_PERIOD_ID", igrDUTY_PERIOD.GetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("DUTY_PERIOD_ID")));
                        IDC_DATA_UPDATE_REQUEST_CANCEL.ExecuteNonQuery();
                        vSTATUS = iConv.ISNull(IDC_DATA_UPDATE_REQUEST_CANCEL.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iConv.ISNull(IDC_DATA_UPDATE_REQUEST_CANCEL.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return;
                        }
                        object mValue;
                        mValue = IDC_DATA_UPDATE_REQUEST_CANCEL.GetCommandParamValue("O_APPROVE_STATUS");
                        igrDUTY_PERIOD.SetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS"), mValue);
                        mValue = IDC_DATA_UPDATE_REQUEST_CANCEL.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                        igrDUTY_PERIOD.SetCellValue(R, igrDUTY_PERIOD.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                    }
                } 
            } 
            igrDUTY_PERIOD.LastConfirmChanges();
            idaDUTY_PERIOD.OraSelectData.AcceptChanges();
            idaDUTY_PERIOD.Refillable = true;
            isSearch_DB(); 
        }

        private void BTN_SELECT_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                idaDUTY_PERIOD.Update();
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                return;
            }
             
            object vDUTY_PERIOD_ID = igrDUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID");
            object vPERSON_NUM = igrDUTY_PERIOD.GetCellValue("PERSON_NUM");
            object vSTART_DATE = igrDUTY_PERIOD.GetCellValue("START_DATE");
            object vEND_DATE = igrDUTY_PERIOD.GetCellValue("END_DATE");
            if (iConv.ISNull(vDUTY_PERIOD_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10209"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            object vDUTY_PERIOD_ID = igrDUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID");
            if (iConv.ISNull(vDUTY_PERIOD_ID) == string.Empty)
            {
                return;
            }

            IDC_GET_DOC_LAST_FLAG.SetCommandParamValue("W_DUTY_PERIOD_ID", igrDUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID"));
            IDC_GET_DOC_LAST_FLAG.ExecuteNonQuery();
            string vLAST_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_FLAG.GetCommandParamValue("O_LAST_FLAG"));
            if (vLAST_FLAG == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            mSave_Flag = false;
            if(iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=사원 정보"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["START_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=시작일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["END_DATE"]) == string.Empty)
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
            //저장전 검증.
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_SELECT_YN", e.Row["SELECT_YN"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_CORP_ID", e.Row["CORP_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_DUTY_ID", e.Row["DUTY_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_START_DATE", e.Row["START_DATE"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_START_TIME", e.Row["START_TIME_H"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_START_TIME_M", e.Row["START_TIME_M"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_END_DATE", e.Row["END_DATE"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_END_TIME", e.Row["END_TIME_H"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_END_TIME_M", e.Row["END_TIME_M"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_DESCRIPTION", e.Row["DESCRIPTION"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_DUTY_PERIOD_ID", e.Row["DUTY_PERIOD_ID"]);
            IDC_VALIDATE_SAVE_P.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_VALIDATE_SAVE_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_VALIDATE_SAVE_P.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            mSave_Flag = true;
        }

        private void idaDUTY_PERIOD_PreDelete(ISPreDeleteEventArgs e)
        {
        
        }

        private void idaDUTY_PERIOD_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if(pBindingManager.DataRow == null)
            {
                isSearch_WorkCalendar(0, igrDUTY_PERIOD.GetCellValue("START_DATE"), igrDUTY_PERIOD.GetCellValue("END_DATE"));
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, 0);
                return;
            } 
            isSearch_WorkCalendar(igrDUTY_PERIOD.GetCellValue("PERSON_ID"), igrDUTY_PERIOD.GetCellValue("START_DATE"), igrDUTY_PERIOD.GetCellValue("END_DATE"));
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, igrDUTY_PERIOD.GetCellValue("DUTY_PERIOD_ID"));
        }

        #endregion

        #region ----- LookUp Event -----

        private void ilaAPPROVE_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildAPPROVE_STATUS.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ildAPPROVE_STATUS.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaSEARCH_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildSEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            ildSEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }
        
        private void ilaDUTY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_CODE.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ildDUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_CODE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        private void ilaDUTY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_CODE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        private void ilaSTART_TIME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", igrDUTY_PERIOD.GetCellValue("START_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", "Y");
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", null);
        }

        private void ilaEND_TIME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", igrDUTY_PERIOD.GetCellValue("END_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", null);
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", "Y");
        }

        #endregion

        private void BTN_GET_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int mRECORD_COUNT = 0;

            if (W_START_DATE.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_START_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (W_END_DATE.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_END_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //if (mRECORD_COUNT != Convert.ToInt32(0))
            //{
            //    //[2011-07-25]
            //    idaOT_HEADER.Cancel();
            //    //기준일자에 대한 연장근무 신청이 이미 존재 합니다. 신청No로 조회해 수정 하십시오!
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10301"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            try
            {
                PB_GET_PERSON.Visible = true;
                idaDUTY_PERIOD.Cancel();
                IDA_INSERT_DATA.Fill();

                int vCountRow = IDA_INSERT_DATA.OraSelectData.Rows.Count;
                int vCountColumn = IDA_INSERT_DATA.OraSelectData.Columns.Count - 2;

                idaDUTY_PERIOD.MoveLast(igrDUTY_PERIOD.Name);
                int vIDX_CURR = igrDUTY_PERIOD.RowIndex;
                if (vIDX_CURR == -1)
                {
                    idaDUTY_PERIOD.Cancel();
                }
                vIDX_CURR = igrDUTY_PERIOD.RowIndex;

                if (vCountRow > 0)
                {
                    igrDUTY_PERIOD.BeginUpdate();
                    for (int vROW = 0; vROW < vCountRow; vROW++)
                    {
                        idaDUTY_PERIOD.AddUnder();
                        for (int vCOL = 0; vCOL < vCountColumn; vCOL++)
                        {
                            igrDUTY_PERIOD.SetCellValue(vROW + (vIDX_CURR + 1), vCOL, IDA_INSERT_DATA.OraSelectData.Rows[vROW][vCOL]);
                        }

                        float vBarFill = ((float)vROW / (float)(vCountRow - 1)) * 100;
                        PB_GET_PERSON.BarFillPercent = vBarFill;
                    }
                    igrDUTY_PERIOD.EndUpdate();
                }
                igrDUTY_PERIOD.CurrentCellMoveTo(0, 0);
                igrDUTY_PERIOD.CurrentCellActivate(0, 0);
                igrDUTY_PERIOD.Focus();

                PB_GET_PERSON.Visible = false;
            }
            catch (System.Exception ex)
            {
                PB_GET_PERSON.Visible = false;

                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e)
        {

            int vIDX_SELECT_FLAG = igrDUTY_PERIOD.GetColumnToIndex("SELECT_YN");
            for (int r = 0; r < igrDUTY_PERIOD.RowCount; r++)
            {
                igrDUTY_PERIOD.SetCellValue(r, vIDX_SELECT_FLAG, V_SELECT_YN.CheckBoxString);
            }
        }

        private void ILA_START_DATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", igrDUTY_PERIOD.GetCellValue("START_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", "Y");
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", null);
        }

        private void ILA_END_DATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERIOD_TIME.SetLookupParamValue("W_WORK_DATE", igrDUTY_PERIOD.GetCellValue("END_DATE"));
            ildPERIOD_TIME.SetLookupParamValue("W_START_YN", null);
            ildPERIOD_TIME.SetLookupParamValue("W_END_YN", "Y");
        }

     
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