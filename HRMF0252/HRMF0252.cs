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
using Syncfusion.GridExcelConverter;
 
using System.ComponentModel;
using System.Text;

using System.Reflection;
using System.Diagnostics;



namespace HRMF0252
{
    public partial class HRMF0252 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISCommonUtil.ISFunction.ISConvert iConvert = new ISFunction.ISConvert();

        ISHR.isCertificatePrint mPrintInfo;
        #endregion;

        #region ----- UpLoad / DownLoad Variables -----

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private ItemImageInfomationFTP mImageFTP;

        private string mFTP_Source_Directory = string.Empty;            // ftp 소스 디렉토리.
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.
        private string mClient_Directory = string.Empty;                // 실제 디렉토리 
        private string mClient_ImageDirectory = string.Empty;           // 클라이언트 이미지 디렉토리.
        private string mFileExtension = ".bmp";                         // 확장자명.

        private bool mIsGetInformationFTP = false;                      // FTP 정보 상태.
        private bool mIsFormLoad = false;                               // NEWMOVE 이벤트 제어.
        private int mStartPage = 1;                                     // 시작 페이지
        
        #endregion;

        #region ----- Constructor -----

        public HRMF0252()
        {
            InitializeComponent();
        }

        public HRMF0252(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;


            mPrintInfo = new ISHR.isCertificatePrint();
            //mPrintInfo = pPrintInfo;
            //mPrintInfo.ISPrinting += ISOnPrint;
        }

        public HRMF0252(Form pMainForm, ISAppInterface pAppInterface, object pJOB_NO)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_CERTI_TYPE.EditValue = pJOB_NO;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SEARCH_DB()
        {
            IDA_CERTIFICATE_PRINT.Cancel();
            IDA_CERTIFICATE_PRINT.Fill();
        }

        #endregion;

        #region ----- Events -----

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
                    //IDA_APPROVED_CERTI.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    //ExcelExport(IGR_APPROVED_CERTI);
                }
            }
        }

        #endregion;

        #region ----- Excel Export -----
        private void ExcelExport(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            GridExcelConverterControl vExport = new GridExcelConverterControl();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xls)|*.xls";
            saveFileDialog.DefaultExt = ".xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ////데이터 테이블을 이용한 export
                //Syncfusion.XlsIO.ExcelEngine vEng = new Syncfusion.XlsIO.ExcelEngine();
                //Syncfusion.XlsIO.IApplication vApp = vEng.Excel;
                //string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();
                //if (vFileExtension == "XLSX")
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel2007;
                //}
                //else
                //{
                //    vApp.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Excel97to2003;
                //}
                //Syncfusion.XlsIO.IWorkbook vWorkbook = vApp.Workbooks.Create(1);
                //Syncfusion.XlsIO.IWorksheet vSheet = vWorkbook.Worksheets[0];
                //foreach(System.Data.DataRow vRow in IDA_MATERIAL_LIST_ALL.CurrentRows)
                //{
                //    vSheet.ImportDataTable(vRow.Table, true, 1, 1, -1, -1);
                //}
                //vWorkbook.SaveAs(saveFileDialog.FileName);
                vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);
                if (MessageBox.Show("Do you wish to open the xls file now?",
                                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }
        #endregion

        #region ----- Assembly Run Methods ----
        private void Print_10( object pEMPLOYE
                                            , object pPERSON_ID
                                            , object PERSON_NUM
                                            , object PERSON_NMAE
                                            , object CERTI_TYPE
                                            , object CERTI_TYPE_ID
                                            , object CERT_TYPE_NAME
                                            , object PERPOSE
                                            , object DESCRIPTION
                                            , object JOINDATE
                                            , object RETIREDATE
                                            , object CORP_ID

                                )
        {
            string vAssemblyId = string.Empty;
            string vAssemblyFileName = string.Empty;
            string vAssemblyFileVersion = string.Empty;

            vAssemblyId = "HRMF0240";
            vAssemblyFileName = "HRMF0240.dll";

            AssmblyRun_Manual(vAssemblyId);

            Assembly vAssembly = Assembly.LoadFrom(vAssemblyFileName);
            Type vType = vAssembly.GetType(vAssembly.GetName().Name + "." + vAssembly.GetName().Name);

            if (vType != null)
            {
                object[] vParam = new object[15];
                vParam[0] = this.MdiParent;
                vParam[1] = isAppInterfaceAdv1.AppInterface;
                vParam[2] = this.mPrintInfo;
                vParam[3] = pEMPLOYE;
                vParam[4] = pPERSON_ID;
                vParam[5] = PERSON_NUM;
                vParam[6] = PERSON_NMAE;
                vParam[7] = CERTI_TYPE;
                vParam[8] = CERTI_TYPE_ID;
                vParam[9] = CERT_TYPE_NAME;
                vParam[10] = PERPOSE;
                vParam[11] = DESCRIPTION;
                vParam[12] = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(JOINDATE));
                vParam[13] = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(RETIREDATE));
                vParam[14] = CORP_ID;


                object vCreateInstance = Activator.CreateInstance(vType, vParam);
                //object vTest = Activator.CreateInstance(vType, vParam);
                //Form vForm = vCreateInstance as Form;
                Office2007Form vForm = vCreateInstance as Office2007Form;
                Point vPoint = new Point(30, 30);
                vForm.Location = vPoint;
                vForm.StartPosition = FormStartPosition.Manual;
                vForm.Text = string.Format("{0}[{1}] - {2}", "Cetificate Print", vAssemblyId, vAssemblyFileVersion);

                vForm.Show(); 
            }
        }

        private void Print_20(object pPERSON_ID
                            , object pCORP_ID
                            , object pPERSON_NUM
                            , object pPERSON_NMAE
                            , object pCERTI_TYPE
                            , object pCERTI_TYPE_ID
                            , object pCERT_TYPE_NAME
                            , object pPERPOSE
                            , object pDESCRIPTION
                            , object pJOIN_DATE
                            , object pRETIRE_DATE
                            )
        {

            string vAssemblyId = string.Empty;
            string vAssemblyFileName = string.Empty;
            string vAssemblyFileVersion = string.Empty;

            vAssemblyId = "HRMF0730";
            vAssemblyFileName = "HRMF0730.dll";

            AssmblyRun_Manual(vAssemblyId);

            Assembly vAssembly = Assembly.LoadFrom(vAssemblyFileName);
            Type vType = vAssembly.GetType(vAssembly.GetName().Name + "." + vAssembly.GetName().Name);

            if (vType != null)
            {
                object[] vParam = new object[15];
                vParam[0] = this.MdiParent;
                vParam[1] = isAppInterfaceAdv1.AppInterface;
                vParam[2] = iConv.ISDecimaltoZero(pCORP_ID);
                vParam[3] = iConv.ISNull(pPERSON_NUM);
                vParam[4] = string.Format("{0:yyyy-MM-dd}", DateTime.Today);
                vParam[5] = iConv.ISNull(((DateTime.Today.Year) -1));
                vParam[6] = iConv.ISNull(pCERT_TYPE_NAME);
                vParam[7] = pCERTI_TYPE_ID;
                vParam[8] = pCERTI_TYPE;
                vParam[9] = iConv.ISNull(pPERSON_NMAE);
                vParam[10] = pPERSON_ID;
                vParam[11] = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(pJOIN_DATE)); 
                vParam[12] = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(pRETIRE_DATE)); 
                vParam[13] = iConv.ISNull(pPERPOSE);
                vParam[14] = iConv.ISNull(pDESCRIPTION);


                object vCreateInstance = Activator.CreateInstance(vType, vParam);
                //object vTest = Activator.CreateInstance(vType, vParam);
                //Form vForm = vCreateInstance as Form;
                Office2007Form vForm = vCreateInstance as Office2007Form;
                //Point vPoint = new Point(30, 30);
                //vForm.Location = vPoint;
                vForm.StartPosition = FormStartPosition.Manual;
                vForm.Text = string.Format("{0}[{1}] - {2}", "Certificate Print", vAssemblyId, vAssemblyFileVersion);

                vForm.Show(); 
            }
        }

        private void AssmblyRun_Manual(object pAssembly_ID)
        {
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vCurrAssemblyFileVersion = string.Empty;

            //[EAPP_ASSEMBLY_INFO_G.MENU_ENTRY_PROCESS_START]
            IDC_MENU_ENTRY_MANUAL_START.SetCommandParamValue("W_ASSEMBLY_ID", pAssembly_ID);
            IDC_MENU_ENTRY_MANUAL_START.ExecuteNonQuery();

            string vREAD_FLAG = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_READ_FLAG"));
            string vUSER_TYPE = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_USER_TYPE"));
            string vPRINT_FLAG = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_PRINT_FLAG"));

            decimal vASSEMBLY_INFO_ID = iConvert.ISDecimaltoZero(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_INFO_ID"));
            string vASSEMBLY_ID = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_ID"));
            string vASSEMBLY_NAME = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_NAME"));
            string vASSEMBLY_FILE_NAME = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_FILE_NAME"));

            string vASSEMBLY_VERSION = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_VERSION"));
            string vDIR_FULL_PATH = iConvert.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_DIR_FULL_PATH"));

            System.IO.FileInfo vFile = new System.IO.FileInfo(vASSEMBLY_FILE_NAME);
            if (vFile.Exists)
            {
                vCurrAssemblyFileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(vASSEMBLY_FILE_NAME).FileVersion;
            }


            //1. Assembly file Name(.dll) 있을겨우만 실행//
            if (vASSEMBLY_FILE_NAME != string.Empty)
            {
                //2. 읽기 권한 있을 경우만 실행 //
                if (vREAD_FLAG == "Y")
                {
                    if (vCurrAssemblyFileVersion != vASSEMBLY_VERSION)
                    {
                        ISFileTransferAdv vFileTransferAdv = new ISFileTransferAdv();

                        vFileTransferAdv.Host = isAppInterfaceAdv1.AppInterface.AppHostInfo.Host;
                        vFileTransferAdv.Port = isAppInterfaceAdv1.AppInterface.AppHostInfo.Port;
                        vFileTransferAdv.UserId = isAppInterfaceAdv1.AppInterface.AppHostInfo.UserId;
                        vFileTransferAdv.Password = isAppInterfaceAdv1.AppInterface.AppHostInfo.Password;
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "N")
                        {
                            vFileTransferAdv.UsePassive = false;
                        }
                        else
                        {
                            vFileTransferAdv.UsePassive = true;
                        } 
                        vFileTransferAdv.SourceDirectory = vDIR_FULL_PATH;
                        vFileTransferAdv.SourceFileName = vASSEMBLY_FILE_NAME;
                        vFileTransferAdv.TargetDirectory = Application.StartupPath;
                        vFileTransferAdv.TargetFileName = "_" + vASSEMBLY_FILE_NAME;

                        if (vFileTransferAdv.Download() == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vASSEMBLY_FILE_NAME);
                                System.IO.File.Move("_" + vASSEMBLY_FILE_NAME, vASSEMBLY_FILE_NAME);
                            }
                            catch
                            {
                                try
                                {
                                    System.IO.FileInfo vFileInfo = new System.IO.FileInfo("_" + vASSEMBLY_FILE_NAME);
                                    if (vFileInfo.Exists == true)
                                    {
                                        try
                                        {
                                            System.IO.File.Delete("_" + vASSEMBLY_FILE_NAME);
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
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10241"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        //report update//
                        ReportUpdate(vASSEMBLY_INFO_ID);
                    } 
                }
            }

            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        //report download//
        private void ReportUpdate(object pAssemblyInfoID)
        {
            string vPathReportFTP = string.Empty;
            string vReportFileName = string.Empty;
            string vReportFileNameTarget = string.Empty;

            try
            {

                IDA_REPORT_INFO_DOWNLOAD.SetSelectParamValue("W_ASSEMBLY_INFO_ID", pAssemblyInfoID);
                IDA_REPORT_INFO_DOWNLOAD.Fill();
                if (IDA_REPORT_INFO_DOWNLOAD.OraSelectData.Rows.Count > 0)
                {
                    ISFileTransferAdv vFileTransferAdv = new ISFileTransferAdv();

                    vFileTransferAdv.Host = isAppInterfaceAdv1.AppInterface.AppHostInfo.Host;
                    vFileTransferAdv.Port = isAppInterfaceAdv1.AppInterface.AppHostInfo.Port;
                    if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive == "N")
                    {
                        vFileTransferAdv.UsePassive = false;
                    }
                    else
                    {
                        vFileTransferAdv.UsePassive = true;
                    }
                    vFileTransferAdv.UserId = isAppInterfaceAdv1.AppInterface.AppHostInfo.UserId;
                    vFileTransferAdv.Password = isAppInterfaceAdv1.AppInterface.AppHostInfo.Password;

                    foreach (System.Data.DataRow vRow in IDA_REPORT_INFO_DOWNLOAD.OraSelectData.Rows)
                    {
                        if (iConvert.ISNull(vRow["REPORT_FILE_NAME"]) != string.Empty)
                        {
                            vReportFileName = iConvert.ISNull(vRow["REPORT_FILE_NAME"]);
                            vReportFileNameTarget = string.Format("_{0}", vReportFileName);
                        }
                        if (iConvert.ISNull(vRow["REPORT_PATH_FTP"]) != string.Empty)
                        {
                            vPathReportFTP = iConvert.ISNull(vRow["REPORT_PATH_FTP"]);
                        }

                        if (vReportFileName != string.Empty && vPathReportFTP != string.Empty)
                        {
                            string vPathReportClient = string.Format("{0}\\{1}", System.Windows.Forms.Application.StartupPath, "Report");
                            System.IO.DirectoryInfo vReport = new System.IO.DirectoryInfo(vPathReportClient);
                            if (vReport.Exists == false) //있으면 True, 없으면 False
                            {
                                vReport.Create();
                            }
                            ////------------------------------------------------------------------------
                            ////[Test Path]
                            ////------------------------------------------------------------------------
                            //string vPathTest = @"K:\00_2_FXE\ERPMain\FXEMain\bin\Debug";
                            //string vPathReportClient = string.Format("{0}\\{1}", vPathTest, "Report");
                            ////------------------------------------------------------------------------

                            vFileTransferAdv.SourceDirectory = vPathReportFTP;
                            vFileTransferAdv.SourceFileName = vReportFileName;
                            vFileTransferAdv.TargetDirectory = vPathReportClient;
                            vFileTransferAdv.TargetFileName = vReportFileNameTarget;

                            string vFullPathReportClient = string.Format("{0}\\{1}", vPathReportClient, vReportFileName);
                            string vFullPathReportTarget = string.Format("{0}\\{1}", vPathReportClient, vReportFileNameTarget);

                            if (vFileTransferAdv.Download() == true)
                            {
                                try
                                {
                                    System.IO.File.Delete(vFullPathReportClient);
                                    System.IO.File.Move(vFullPathReportTarget, vFullPathReportClient);
                                }
                                catch
                                {
                                    try
                                    {
                                        System.IO.FileInfo vFileInfo = new System.IO.FileInfo(vFullPathReportTarget);
                                        if (vFileInfo.Exists == true)
                                        {
                                            System.IO.File.Delete(vFullPathReportTarget);
                                        }
                                    }
                                    catch
                                    {
                                        //
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
            }
        }

        #endregion;

        private void HRMF0252_Load(object sender, EventArgs e)
        {
            IDA_CERTIFICATE_PRINT.FillSchema();
        }

        private void HRMF0252_Shown(object sender, EventArgs e)
        {
            //DEFAULT Date SETTING
            iSTART_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            iEND_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);
            
            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        

            W_CORP_NAME_0.BringToFront();

        }

        #region ----- Form Event ------

        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e)
        {
     
        }

        #endregion

        private void Default_Setting()
        {
            IGR_CERTIFICATE_PRINT.SetCellValue("PRINT_DATE", DateTime.Today );
        }
        
        #region ----- Lookup Event ------

        private void ILA_CERTIFICATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CERTIFICATE_W.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ILD_CERTIFICATE_W.SetLookupParamValue("W_WHERE", "HC.VALUE3 = 'Y'");
            ILD_CERTIFICATE_W.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

     
        private void ILA_SEARCH_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SEARCH_TYPE.SetLookupParamValue("W_GROUP_CODE", "SEARCH_TYPE");
            ILD_SEARCH_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }





        #endregion

        private void IGR_APPROVED_CERTI_CellDoubleClick(object pSender)
        {
            
            object vPERSON_ID = IGR_CERTIFICATE_PRINT.GetCellValue("PERSON_ID");
            object vCERTI_TYPE_ID = IGR_CERTIFICATE_PRINT.GetCellValue("CERT_TYPE_ID");
            object vCERTI_TYPE_CODE = IGR_CERTIFICATE_PRINT.GetCellValue("CERT_TYPE");
            object vCERTI_TYPE = IGR_CERTIFICATE_PRINT.GetCellValue("CERT_CATEGORY");
            object vCERT_TYPE_NAME = IGR_CERTIFICATE_PRINT.GetCellValue("CERT_TYPE_NAME");
            object vPERSON_NUM = IGR_CERTIFICATE_PRINT.GetCellValue("PERSON_NUM");
            object vPERSON_NMAE = IGR_CERTIFICATE_PRINT.GetCellValue("NAME");
            object vPERPOSE = IGR_CERTIFICATE_PRINT.GetCellValue("CERT_PRINT_PERPOSE");
            object vDESCRIPTION = IGR_CERTIFICATE_PRINT.GetCellValue("DESCRIPTION");
            object vEMPLOYE = IGR_CERTIFICATE_PRINT.GetCellValue("EMPLOYE_TYPE");
            object vCORP_ID = IGR_CERTIFICATE_PRINT.GetCellValue("CORP_ID");
            object vJOIN_DATE = IGR_CERTIFICATE_PRINT.GetCellValue("JOIN_DATE");
            object vRETIRE_DATE = IGR_CERTIFICATE_PRINT.GetCellValue("RETIRE_DATE");
            if (iConv.ISNull(vPERSON_ID) != string.Empty && iConv.ISNull(vCERTI_TYPE_ID) != string.Empty)
            {
                if(iConv.ISNull(vCERTI_TYPE) == "10")
                {
                    Print_10(vEMPLOYE
                            , vPERSON_ID
                            , vPERSON_NUM
                            , vPERSON_NMAE
                            , vCERTI_TYPE_CODE
                            , vCERTI_TYPE_ID
                            , vCERT_TYPE_NAME
                            , vPERPOSE
                            , vDESCRIPTION
                            , vJOIN_DATE
                            , vRETIRE_DATE
                            , vCORP_ID
                            );
                }
                else
                {
                    Print_20(vPERSON_ID
                            , vCORP_ID
                            , vPERSON_NUM
                            , vPERSON_NMAE
                            , vCERTI_TYPE_CODE
                            , vCERTI_TYPE_ID
                            , vCERT_TYPE_NAME
                            , vPERPOSE
                            , vDESCRIPTION
                            , vJOIN_DATE
                            , vRETIRE_DATE
                            );
                }
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10033"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        
    }
        #region ----- User Make Class -----

        public class ItemImageInfomationFTP
        {
            #region ----- Variables -----

            private string mHost = string.Empty;
            private string mPort = "21";
            private string mUserID = string.Empty;
            private string mPassword = string.Empty;
            private string mPassive_Flag = "N";

            #endregion;

            #region ----- Constructor -----

            public ItemImageInfomationFTP()
            {
            }

            public ItemImageInfomationFTP(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag)
            {
                mHost = pHost;
                mPort = pPort;
                mUserID = pUserID;
                mPassword = pPassword;
                mPassive_Flag = pPassive_Flag;
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

            #endregion;
        }

        #endregion;
    }
}