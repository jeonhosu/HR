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
using System.IO;

namespace HRMF0761
{
    public partial class HRMF0761 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.

        #endregion;

        #region ----- Constructor ----- 

        public HRMF0761()
        {
            InitializeComponent();
        }

        public HRMF0761(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- User Make Methods ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void DefaultDate()
        {
            if (DateTime.Today.Month <= 2)
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(iDate.ISDate_Add(string.Format("{0}-01-01", DateTime.Today.Year), -1));
            }
            else
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            }
        }

        private DateTime GetDateTime()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void SEARCH_DB()
        {
            string vMessage = string.Empty;
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_STD_YYYYMM.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            try
            {
                string vPERSON_NUM = iConv.ISNull(IGR_PERSON.GetCellValue("PERSON_NUM"));
                int vIDX_Col = IGR_PERSON.GetColumnToIndex("PERSON_NUM");

                IDA_PERSON.Fill();
                if (IGR_PERSON.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_PERSON.RowCount; vRow++)
                    {
                        if (vPERSON_NUM == iConv.ISNull(IGR_PERSON.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_PERSON.CurrentCellActivate(vRow, 0);
                            IGR_PERSON.CurrentCellMoveTo(vRow, 0);
                            IGR_PERSON.Focus();
                            return;
                        }
                    }
                }
                IGR_PERSON.Focus();
            }
            catch (System.Exception ex)
            {
                vMessage = string.Format("Adapter Fill Error\n{0}", ex.Message);
                MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
         
        private void SetCommon(object pGROUP_CODE, object pENABLED_FLAG_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pENABLED_FLAG_YN);
        }
         
        #endregion;

        #region ---- NTS Show -----

        private void Show_NTS_Reader(int pROW_INDEX, object pYYYY, object pPERSON_ID, object pPERSON_NUM, object pNAME, object pYESONE_TYPE
                                    , string pFTP_Filename)
        {
            if (iConv.ISNull(pYYYY) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(pPERSON_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(pYESONE_TYPE) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10155"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            //기간 마감 여부.
            IDC_CLOSING_CHECK_P.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS != "S")
            {
                if(vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            DialogResult vdlgResult;
            NTS_Reader vNTS_Reader = new NTS_Reader(isAppInterfaceAdv1.AppInterface, pYYYY, pNAME, pPERSON_NUM);
            vdlgResult = vNTS_Reader.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                MessageBoxAdv.Show("error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            
            string vPDF_Filename = iConv.ISNull(vNTS_Reader.PDF_Filename);
            object vPDF_PWD = vNTS_Reader.PDF_PWD;
            string vStrBuf = vNTS_Reader.PDF_StrBuf; 
            

            string vUSER_FILENAME = Path.GetFileName(vPDF_Filename);
            string vEXTENSION_NAME = Path.GetExtension(vPDF_Filename).ToUpper();

            IDC_IMPORT_PDF.DataTransaction = isDataTransaction1;
            IDC_EXEC_YESONE_PDF_FILE.DataTransaction = isDataTransaction1;
            isDataTransaction1.BeginTran();

            IDC_IMPORT_PDF.SetCommandParamValue("P_YYYY", pYYYY);
            IDC_IMPORT_PDF.SetCommandParamValue("P_PERSON_ID", pPERSON_ID);
            IDC_IMPORT_PDF.SetCommandParamValue("P_XMLDATA", vStrBuf); 
            IDC_IMPORT_PDF.SetCommandParamValue("P_YESONE_TYPE", pYESONE_TYPE);
            IDC_IMPORT_PDF.SetCommandParamValue("P_PWD", vPDF_PWD);
            IDC_IMPORT_PDF.SetCommandParamValue("P_USER_FILENAME", vUSER_FILENAME);
            IDC_IMPORT_PDF.SetCommandParamValue("P_EXTENSION_NAME", vEXTENSION_NAME);
            IDC_IMPORT_PDF.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_IMPORT_PDF.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_IMPORT_PDF.GetCommandParamValue("O_MESSAGE"));
            string vFTP_FILENAME = iConv.ISNull(IDC_IMPORT_PDF.GetCommandParamValue("O_FTP_FILENAME"));
            if (vSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                IDC_IMPORT_PDF.DataTransaction = null;
                IDC_EXEC_YESONE_PDF_FILE.DataTransaction = null;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
                return;
            }

            if (pFTP_Filename != string.Empty)
            {
                //기존 ftp 파일 삭제//
                mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;  //삭제는 소스에 설정해야 삭제됨.
                mFileTransfer.SourceFileName = pFTP_Filename;
                mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
                mFileTransfer.TargetFileName = pFTP_Filename;

                bool IsDelete = mFileTransfer.Delete();
            }

            if (UpLoadFile(vPDF_Filename, vFTP_FILENAME) == false)
            {
                isDataTransaction1.RollBack();
                IDC_IMPORT_PDF.DataTransaction = null;
                IDC_EXEC_YESONE_PDF_FILE.DataTransaction = null;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents(); 
                return;
            }
             
            //FLAG UPDATE
            IDC_EXEC_YESONE_PDF_FILE.SetCommandParamValue("P_YYYY", pYYYY);
            IDC_EXEC_YESONE_PDF_FILE.SetCommandParamValue("P_PERSON_ID", pPERSON_ID);
            IDC_EXEC_YESONE_PDF_FILE.SetCommandParamValue("P_YESONE_TYPE", pYESONE_TYPE);
            IDC_EXEC_YESONE_PDF_FILE.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_EXEC_YESONE_PDF_FILE.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_EXEC_YESONE_PDF_FILE.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                IDC_IMPORT_PDF.DataTransaction = null;
                IDC_EXEC_YESONE_PDF_FILE.DataTransaction = null;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            isDataTransaction1.Commit();
            IDC_IMPORT_PDF.DataTransaction = null;
            IDC_EXEC_YESONE_PDF_FILE.DataTransaction = null;

            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("PDF_FILE_FLAG"), "Y");
            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("USER_FILENAME"), vUSER_FILENAME);
            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("FTP_FILENAME"), vFTP_FILENAME);

            IDA_PERSON.OraSelectData.AcceptChanges();
            IDA_PERSON.Refillable = true;
            IGR_PERSON.LastConfirmChanges();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents(); 
        }

        #endregion

        #region ----- FTP Infomation -----
        //ftp 접속정보 및 환경 정보 설정 
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
             
            mFTP_Connect_Status = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "YESONE_PDF");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
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
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mFTP_Info.Host == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
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
                System.Windows.Forms.Cursor.Current = Cursors.Default;
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
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
         
        //ftp에 file upload 처리 
        private bool UpLoadFile(string pUser_FileName, string pFTP_FileName)
        {
            bool isUpload = false;

            if (mFTP_Connect_Status == false)
            {
                MessageBoxAdv.Show("FTP Server Connect Fail. Check FTP Server", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                return isUpload;
            }

            if (iConv.ISNull(pFTP_FileName) != string.Empty)
            {
                //1. 사용자 선택 파일 
                string vSelectFullPath = pUser_FileName;
                string vSelectDirectoryPath = Path.GetDirectoryName(pUser_FileName);

                string vFileName = Path.GetFileName(pUser_FileName);
                string vFileExtension = Path.GetExtension(pUser_FileName).ToUpper();

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                //4. 파일 업로드
                try
                {
                    int vArryCount = pFTP_FileName.Length;
                    mFileTransfer.ShowProgress = true;      //진행바 보이기 

                    //업로드 환경 설정 
                    mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                    mFileTransfer.SourceFileName = vFileName;
                    mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
                    mFileTransfer.TargetFileName = iConv.ISNull(pFTP_FileName);

                    bool isUpLoad = mFileTransfer.Upload();

                    if (isUpLoad == true)
                    {
                        isUpload = true;
                    }
                    else
                    {
                        isUpload = false;
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return isUpload;
                    }
                }
                catch (Exception Ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    return isUpload;
                } 
            }
            return isUpload;
        }

        #endregion;

        #region ----- file Download Methods -----
        
        //ftp file download 처리 
        private bool DownLoadFile(string pSAVE_FileName, string pFTP_FILE_NAME)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
              
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

        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(int pROW_INDEX, object pYEAR_YYYY, object pPERSON_ID, object pYESONE_TYPE, string pFTP_FILE_NAME)
        {
            bool IsDelete = false; 

            if (iConv.ISNull(pYEAR_YYYY) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return IsDelete;
            } 

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            //기간 마감 여부.
            IDC_CLOSING_CHECK_P.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS != "S")
            {
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDelete;
            } 

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return IsDelete;
            }

            if (pFTP_FILE_NAME == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            //transaction 이용하기 위해 설정
            IDC_DELETE_YESONE_PDF_FILE.DataTransaction = isDataTransaction1;
            isDataTransaction1.BeginTran();

            //2. 파일 삭제 
            IDC_DELETE_YESONE_PDF_FILE.SetCommandParamValue("P_YYYY", pYEAR_YYYY);
            IDC_DELETE_YESONE_PDF_FILE.SetCommandParamValue("P_PERSON_ID", pPERSON_ID);
            IDC_DELETE_YESONE_PDF_FILE.SetCommandParamValue("P_YESONE_TYPE", pYESONE_TYPE); 
            IDC_DELETE_YESONE_PDF_FILE.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_DELETE_YESONE_PDF_FILE.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_DELETE_YESONE_PDF_FILE.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_YESONE_PDF_FILE.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false;
                isDataTransaction1.RollBack();
                //Transaction 해제.
                IDC_DELETE_YESONE_PDF_FILE.DataTransaction = null;
                
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return IsDelete;
            }

            //3. 실제 삭제  
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;  //삭제는 소스에 설정해야 삭제됨.
            mFileTransfer.SourceFileName = pFTP_FILE_NAME;
            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.TargetFileName = pFTP_FILE_NAME;

            IsDelete = mFileTransfer.Delete();
            if (IsDelete == false)
            {
                isDataTransaction1.RollBack();
                //Transaction 해제.
                IDC_DELETE_YESONE_PDF_FILE.DataTransaction = null;
                
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                return IsDelete;
            }
            isDataTransaction1.Commit();
            //Transaction 해제.
            IDC_DELETE_YESONE_PDF_FILE.DataTransaction = null;

            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("PDF_FILE_FLAG"), "N");
            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("USER_FILENAME"), string.Empty);
            IGR_PERSON.SetCellValue(pROW_INDEX, IGR_PERSON.GetColumnToIndex("FTP_FILENAME"), string.Empty);

            IDA_PERSON.OraSelectData.AcceptChanges();
            IDA_PERSON.Refillable = true;
            IGR_PERSON.LastConfirmChanges();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return IsDelete;
        }

        #endregion;

        #region ----- Get FTP File Name -----
        //ftp file 정보 읽어오기 
        //private void Get_FTP_FileInfo(object pSource_ID)
        //{
        //    IDC_GET_DOC_ATTACHMENT_INFO_P.SetCommandParamValue("P_SOURCE_CATEGORY", vSOURCE_CATEGORY);
        //    IDC_GET_DOC_ATTACHMENT_INFO_P.SetCommandParamValue("P_SOURCE_ID", pSource_ID);
        //    IDC_GET_DOC_ATTACHMENT_INFO_P.ExecuteNonQuery();
        //    O_DOC_ATTACHMENT_ID.EditValue = IDC_GET_DOC_ATTACHMENT_INFO_P.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
        //    O_FTP_FILE_NAME.EditValue = IDC_GET_DOC_ATTACHMENT_INFO_P.GetCommandParamValue("O_FTP_FILE_NAME");
        //}

        #endregion

        #region ----- is View file Method -----
         
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

        #region ----- Main Button Events -----

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
                    if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.AddOver();
                        IGR_FAMILY.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.AddUnder();
                        IGR_FAMILY.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_PERSON.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- This Form Events -----

        private void HRMF0761_Load(object sender, EventArgs e)
        {
            W_CORP_NAME.BringToFront();
            BTN_UPLOAD_ALL.BringToFront();
            BTN_DEL_ALL.BringToFront();
            BTN_VIEW_ALL.BringToFront();

            IDA_PERSON.FillSchema(); 
        }

        private void HRMF0761_Shown(object sender, EventArgs e)
        {
            DefaultDate();
            DefaultCorporation();
            Set_FTP_Info();
        }

        private void V_NAME_KeyDown(object pSender, KeyEventArgs e)
        {
            if(e.Modifiers == Keys.Shift) 
            {
                if (e.KeyCode == Keys.F11)
                {
                    MessageBoxAdv.Show(string.Format("Ftp IP={0} \r\n Port={1} \r\n Folder={2} \r\n ID={3} \r\n Passive={4}"
                                                    , mFTP_Info.Host, mFTP_Info.Port, mFTP_Info.FTP_Folder, mFTP_Info.UserID, mFTP_Info.Passive_Flag), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } 
            }
        }

        private void BTN_UPLOAD_ALL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_NTS_Reader(IGR_PERSON.RowIndex
                            , IGR_PERSON.GetCellValue("ADJUST_YYYY")
                            , IGR_PERSON.GetCellValue("PERSON_ID")
                            , IGR_PERSON.GetCellValue("PERSON_NUM")
                            , IGR_PERSON.GetCellValue("NAME")
                            , "ALL"
                            , iConv.ISNull(IGR_PERSON.GetCellValue("FTP_FILENAME")));
        }

        private void BTN_DEL_ALL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DeleteFile(IGR_PERSON.RowIndex
                    , IGR_PERSON.GetCellValue("ADJUST_YYYY")
                    , IGR_PERSON.GetCellValue("PERSON_ID")
                    , "ALL"
                    , iConv.ISNull(IGR_PERSON.GetCellValue("FTP_FILENAME"))
                    );
        }

        private void BTN_VIEW_ALL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vClientFileName = iConv.ISNull(IGR_PERSON.GetCellValue("USER_FILENAME"));
            string vFTPFileName = iConv.ISNull(IGR_PERSON.GetCellValue("FTP_FILENAME"));
            vClientFileName = isDownload(vClientFileName, vFTPFileName);
            if (vClientFileName == string.Empty)
            {
                return;
            } 
            System.Diagnostics.Process.Start(vClientFileName);
        }

        private void BTN_YESONE_SETUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            string vClientFileName = "YesoneAPISetup.exe";
            string vFTPFileName = "YesoneAPISetup.exe";
            vClientFileName = isDownload(vClientFileName, vFTPFileName);
            if (vClientFileName == string.Empty)
            {
                return;
            }
            System.Diagnostics.Process.Start(vClientFileName);
        }

        private void BTN_EXEC_YEAR_ADJUST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vPERSON_ID = IGR_PERSON.GetCellValue("PERSON_ID"); 
            object vADJUST_DATE = IGR_PERSON.GetCellValue("ADJUST_DATE_TO");
            if(iConv.ISNull(vADJUST_DATE) == String.Empty)
            {
                vADJUST_DATE = iDate.ISMonth_Last(iDate.ISGetDate(W_STD_YYYYMM.EditValue));
            }
            IDC_USER_CAP_R.SetCommandParamValue("W_START_DATE", vADJUST_DATE);
            IDC_USER_CAP_R.SetCommandParamValue("W_END_DATE", vADJUST_DATE);
            IDC_USER_CAP_R.SetCommandParamValue("W_MODULE_CODE", "50");
            IDC_USER_CAP_R.SetCommandParamValue("W_PERSON_ID", isAppInterfaceAdv1.AppInterface.PersonId);
            IDC_USER_CAP_R.ExecuteNonQuery();
            String vCAP_LEVEL  = iConv.ISNull(IDC_USER_CAP_R.GetCommandParamValue("O_CAP_LEVEL"));

            if (vCAP_LEVEL != "C")
            { 
                if (iConv.ISNull(vPERSON_ID) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Questioin", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                string mSTATUS = "F";
                string mMESSAGE = String.Empty;

                //기간 마감 여부.
                IDC_CLOSING_CHECK_20_P.ExecuteNonQuery();
                mSTATUS = iConv.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_STATUS"));
                mMESSAGE = iConv.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_MESSAGE"));
                if (mSTATUS != "S")
                {
                    if (mMESSAGE != String.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();
                 
                object vCORP_ID = IGR_PERSON.GetCellValue("CORP_ID");
                object vADJUST_YYYY = IGR_PERSON.GetCellValue("ADJUST_YYYY");
                object vPERSON_NUM = IGR_PERSON.GetCellValue("PERSON_NUM");
                object vNAME = IGR_PERSON.GetCellValue("NAME");
                 
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_SELECT_YN", "Y");
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_CORP_ID", vCORP_ID);
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_YYYY", vADJUST_YYYY);
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_PERSON_ID", vPERSON_ID);
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_PERSON_NUM", vPERSON_NUM);
                IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_NAME", vNAME);
                IDC_EXEC_YEAR_ADJUST.ExecuteNonQuery();
                mSTATUS = iConv.ISNull(IDC_EXEC_YEAR_ADJUST.GetCommandParamValue("O_STATUS"));
                mMESSAGE = iConv.ISNull(IDC_EXEC_YEAR_ADJUST.GetCommandParamValue("O_MESSAGE"));
                if (IDC_EXEC_YEAR_ADJUST.ExcuteError || mSTATUS == "F")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (mMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }  

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents(); 
            }
            else
            { 
                DialogResult vdlgResult;
                HRMF0761_TRANS vHRMF0761_TRANS = new HRMF0761_TRANS(this.MdiParent, isAppInterfaceAdv1.AppInterface, "EXEC"
                                                                   , W_STD_YYYYMM.EditValue
                                                                   , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                   , W_OPERATING_UNIT_NAME.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                                   , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                   , W_FLOOR_NAME.EditValue, W_FLOOR_ID.EditValue
                                                                   , W_YEAR_EMPLOYE_TYPE_DESC.EditValue, W_YEAR_EMPLOYE_TYPE.EditValue
                                                                   , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
                vdlgResult = vHRMF0761_TRANS.ShowDialog();
                if (vdlgResult == DialogResult.Cancel)
                {
                    return;
                }
            }
            SEARCH_DB();
        }

        private void BTN_CANCEL_YEAR_ADJUST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vPERSON_ID = IGR_PERSON.GetCellValue("PERSON_ID");
            object vADJUST_DATE = IGR_PERSON.GetCellValue("ADJUST_DATE_TO");
            if (iConv.ISNull(vADJUST_DATE) == String.Empty)
            {
                vADJUST_DATE = iDate.ISMonth_Last(iDate.ISGetDate(W_STD_YYYYMM.EditValue));
            }
            IDC_USER_CAP_R.SetCommandParamValue("W_START_DATE", vADJUST_DATE);
            IDC_USER_CAP_R.SetCommandParamValue("W_END_DATE", vADJUST_DATE);
            IDC_USER_CAP_R.SetCommandParamValue("W_MODULE_CODE", "50");
            IDC_USER_CAP_R.SetCommandParamValue("W_PERSON_ID", isAppInterfaceAdv1.AppInterface.PersonId);
            IDC_USER_CAP_R.ExecuteNonQuery();
            String vCAP_LEVEL = iConv.ISNull(IDC_USER_CAP_R.GetCommandParamValue("O_CAP_LEVEL"));

            if (vCAP_LEVEL != "C")
            {
                if (iConv.ISNull(vPERSON_ID) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }


                if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Questioin", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                string mSTATUS = "F";
                string mMESSAGE = String.Empty;

                //기간 마감 여부.
                IDC_CLOSING_CHECK_20_P.ExecuteNonQuery();
                mSTATUS = iConv.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_STATUS"));
                mMESSAGE = iConv.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_MESSAGE"));
                if (mSTATUS != "S")
                {
                    if (mMESSAGE != String.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents(); 
                
                object vCORP_ID = IGR_PERSON.GetCellValue("CORP_ID");
                object vADJUST_YYYY = IGR_PERSON.GetCellValue("ADJUST_YYYY");
                object vPERSON_NUM = IGR_PERSON.GetCellValue("PERSON_NUM");
                object vNAME = IGR_PERSON.GetCellValue("NAME");
                    
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_SELECT_YN", "Y");
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_CORP_ID", vCORP_ID);
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_YYYY", vADJUST_YYYY);
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_PERSON_ID", vPERSON_ID);
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_PERSON_NUM", vPERSON_NUM);
                IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_NAME", vNAME);

                IDC_CANCEL_YEAR_ADJUST.ExecuteNonQuery();
                mSTATUS = iConv.ISNull(IDC_CANCEL_YEAR_ADJUST.GetCommandParamValue("O_STATUS"));
                mMESSAGE = iConv.ISNull(IDC_CANCEL_YEAR_ADJUST.GetCommandParamValue("O_MESSAGE"));
                if (IDC_EXEC_YEAR_ADJUST.ExcuteError || mSTATUS == "F")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    if (mMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                } 

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }
            else
            {
                DialogResult vdlgResult;
                HRMF0761_TRANS vHRMF0761_TRANS = new HRMF0761_TRANS(this.MdiParent, isAppInterfaceAdv1.AppInterface, "CANCEL"
                                                                   , W_STD_YYYYMM.EditValue
                                                                   , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                   , W_OPERATING_UNIT_NAME.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                                   , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                   , W_FLOOR_NAME.EditValue, W_FLOOR_ID.EditValue
                                                                   , W_YEAR_EMPLOYE_TYPE_DESC.EditValue, W_YEAR_EMPLOYE_TYPE.EditValue
                                                                   , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
                vdlgResult = vHRMF0761_TRANS.ShowDialog();
                if (vdlgResult == DialogResult.Cancel)
                {
                    return;
                }
            }
            SEARCH_DB();
        }

        private void IGR_FAMILY_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == IGR_FAMILY.GetColumnToIndex("REPRE_NUM"))
            {
                object vRepre_Num;
                vRepre_Num = e.NewValue;
                if (iConv.ISNull(vRepre_Num) == string.Empty)
                {
                    return;
                }
                if (FAMILY_REPRE_NUM_CHECK(vRepre_Num) == "N".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (iConv.ISNull(IGR_FAMILY.GetCellValue("BIRTHDAY")) == string.Empty)
                {
                    IGR_FAMILY.SetCellValue("BIRTHDAY", BIRTHDAY(vRepre_Num));
                }

                if (iConv.ISNull(IGR_FAMILY.GetCellValue("BIRTHDAY_TYPE")) == string.Empty)
                {
                    // 음양구분.
                    IDC_COMMON_W.SetCommandParamValue("W_GROUP_CODE", "BIRTHDAY_TYPE");
                    IDC_COMMON_W.SetCommandParamValue("W_WHERE", " 1 = 1 ");
                    IDC_COMMON_W.ExecuteNonQuery();
                    IGR_FAMILY.SetCellValue("BIRTHDAY_TYPE_NAME", IDC_COMMON_W.GetCommandParamValue("O_CODE_NAME"));
                    IGR_FAMILY.SetCellValue("BIRTHDAY_TYPE", IDC_COMMON_W.GetCommandParamValue("O_CODE"));
                }
            }
        }

        private DateTime BIRTHDAY(object pREPRE_NUM)
        {
            DateTime mBIRTHDAY;

            string mSex_Type = pREPRE_NUM.ToString().Replace("-", "").Substring(6, 1);
            if (mSex_Type == "1".ToString() || mSex_Type == "2".ToString() || mSex_Type == "5".ToString() || mSex_Type == "6".ToString())
            {
                mBIRTHDAY = DateTime.Parse("19" + pREPRE_NUM.ToString().Substring(0, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(2, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(4, 2));
            }
            else
            {
                mBIRTHDAY = DateTime.Parse("20" + pREPRE_NUM.ToString().Substring(0, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(2, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(4, 2));
            }
            return mBIRTHDAY;
        }

        private string FAMILY_REPRE_NUM_CHECK(object pREPRE_NUM)
        {
            string isReturnValue = "N".ToString();
            if (iConv.ISNull(pREPRE_NUM) == string.Empty)
            {
                return isReturnValue;
            }
            
            // 전호수 주석 : '-' 입력 체크 안함. 단, DB에서 자릿수 검증후 '-' 자동 입력 처리.
            //if (iedREPRE_NUM.EditValue.ToString().IndexOf("-") == -1)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return isReturnValue;
            //}

            IDC_REPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pREPRE_NUM);
            IDC_REPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = IDC_REPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            return isReturnValue;
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ilaYYYYMM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISDate_Month_Add(iDate.ISGetDate(), 4));
        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_EMPLOYE_TYPE", "Y");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("FLOOR", "Y");
        }

        private void ILA_F_RELATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("RELATION", "Y");
        }

        private void ILA_F_YEAR_DISABILITY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_DISABILITY", "Y");
        }

        private void ILA_F_BIRTHDAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("BIRTHDAY_TYPE", "Y");
        }

        private void ILA_F_END_SCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("END_SCH", "Y");
        }

        #endregion

        #region ----- Adapter Event -----


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