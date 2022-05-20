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


namespace HRMF0205
{
    public partial class HRMF0205 : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        #endregion;

        #region ----- Constructor -----

        public HRMF0205(Form pMainForm, ISAppInterface pAppInterface)
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
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void SEARCH_DB()
        {
            IDA_CERTIFICATE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_CERTIFICATE.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_CERTIFICATE.Fill();
        }

        private void isOnPrinting(DateTime pPrint_Date, string pPrint_num)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }

            if (W_STD_DATE.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_DATE.Focus();
                return;
            }

            ISHR.isCertificatePrint vCertificatePrint = new ISHR.isCertificatePrint();
            vCertificatePrint.FormID = this.Name;
            vCertificatePrint.Corp_ID = Convert.ToInt32( W_CORP_ID.EditValue);
            vCertificatePrint.Print_Num = pPrint_num;
            vCertificatePrint.Print_Date = pPrint_Date;
            if (pPrint_num != null)
            {
                vCertificatePrint.Cert_Type_Name = igrCERTIFICATE.GetCellValue("CERT_TYPE_NAME").ToString();
                vCertificatePrint.Cert_Type_ID = Convert.ToInt32(igrCERTIFICATE.GetCellValue("CERT_TYPE_ID"));
                vCertificatePrint.Name = igrCERTIFICATE.GetCellValue("NAME").ToString();
                vCertificatePrint.Person_ID = Convert.ToInt32(igrCERTIFICATE.GetCellValue("PERSON_ID"));
                vCertificatePrint.Join_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("JOIN_DATE"));
                vCertificatePrint.Retire_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("RETIRE_DATE"));
                vCertificatePrint.Description = igrCERTIFICATE.GetCellValue("DESCRIPTION").ToString();
                vCertificatePrint.Send_Org = igrCERTIFICATE.GetCellValue("SEND_ORG").ToString();
                vCertificatePrint.Print_Count = Convert.ToInt32(igrCERTIFICATE.GetCellValue("PRINT_COUNT"));
            }
            
            vCertificatePrint.ISPrinted += ISPrinted;
            ISAppInterface vAppInterface = new ISAppInterface();
            vAppInterface = isAppInterfaceAdv1.AppInterface;
            Form vHRMF0205_PRINT = new HRMF0205_PRINT(this.MdiParent, vCertificatePrint, vAppInterface);
            vCertificatePrint.isPrintingEvent(this.Name);
            vHRMF0205_PRINT.Show();
        }

        // 증명서 관리 폼의 Grid 부분에서 더블클릭 시 해당 내용이 프린트 폼에 표시되며 활성화되는 부분
        private void gridClickPrinting(DateTime dPrint_Date, string sPrint_num, object bCertTypeID, object bCertTypeName, object bName, object bPersonID,
                                       DateTime dJoinDate, DateTime dRetireDate, object bDescription, object bSendOrg/*, int nRetireDateCnt*/)
        {
            ISHR.isCertificatePrint vCertificatePrint = new ISHR.isCertificatePrint();
            vCertificatePrint.FormID = this.Name;
            vCertificatePrint.Corp_ID = Convert.ToInt32(W_CORP_ID.EditValue);
            vCertificatePrint.Print_Num = sPrint_num;
            vCertificatePrint.Print_Date = dPrint_Date;

            if (sPrint_num != null)
            {
                vCertificatePrint.Cert_Type_Name = bCertTypeName.ToString();
                vCertificatePrint.Cert_Type_ID = Convert.ToInt32(bCertTypeID);
                vCertificatePrint.Name = bName.ToString();
                vCertificatePrint.Person_ID = Convert.ToInt32(bPersonID);
                vCertificatePrint.Join_Date = dJoinDate;
                //if (nRetireDateCnt == 1) //퇴직일자가 null이 아니면 날짜를 넘겨줌
                //{
                    vCertificatePrint.Retire_Date = dRetireDate;
                //}
                vCertificatePrint.Description = bDescription.ToString();
                vCertificatePrint.Send_Org = bSendOrg.ToString();
                vCertificatePrint.Print_Count = Convert.ToInt32(igrCERTIFICATE.GetCellValue("PRINT_COUNT"));
            }

            vCertificatePrint.ISPrinted += ISPrinted;
            ISAppInterface vAppInterface = new ISAppInterface();
            vAppInterface = isAppInterfaceAdv1.AppInterface;
            Form vHRMF0205_PRINT = new HRMF0205_PRINT(this.MdiParent, vCertificatePrint, vAppInterface);
            vCertificatePrint.isPrintingEvent(this.Name);
            vHRMF0205_PRINT.Show();
        }

        #endregion;

        #region -----isAppInterfaceAdv1_AppMainButtonClick Events -----
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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    AssmblyRun_Manual("HRMF0240", W_CORP_ID.EditValue, "", ""); 
                }                
            }
        }

        private void ISPrinted(string pFormID)
        {
            SEARCH_DB();
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0205_Load(object sender, EventArgs e)
        {
            W_STD_DATE.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }
        #endregion

        #region ----- Lookup Event -----
        private void ilaCERT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 10 ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", W_STD_DATE.EditValue);
        }

        private void igrCERTIFICATE_DoubleClick(object sender, EventArgs e)
        {
            DateTime dPrint_Date = DateTime.Today;
            string sPrint_num = null;

            if (igrCERTIFICATE.RowCount > 0)
            {
                dPrint_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("PRINT_DATE"));
                sPrint_num = igrCERTIFICATE.GetCellValue("PRINT_NUM").ToString();

            }
            isOnPrinting(dPrint_Date, sPrint_num);
        }
        #endregion

        #region ----- Convert Date Methods ----

        private System.DateTime ConvertDate(object pObject)
        {
            System.DateTime vDateTime = new DateTime();

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        vDateTime = (System.DateTime)pObject;                        
                    }
                }
            }
            catch
            {

            }

            return vDateTime;
        }

        #endregion; 

        private void igrCERTIFICATE_CellDoubleClick(object pSender)
        {
            if (igrCERTIFICATE.Row < 1)
                return;

            object vPRINT_NUM = igrCERTIFICATE.GetCellValue("PRINT_NUM");     //발급번호.
            object vPRINT_DATE= igrCERTIFICATE.GetCellValue("PRINT_DATE");     //인쇄일자.

            AssmblyRun_Manual("HRMF0240", W_CORP_ID.EditValue, vPRINT_NUM, vPRINT_DATE); 

            //DateTime dPrint_Date = DateTime.Today; //발급일자(기준일자)      
            //string sPrint_num = null;              //발급번호

            //object bCertTypeID = null;
            //object bCertTypeName = null;
            //object bName = null;
            //object bPersonID = null;
            //object bDescription = null;
            //object bSendOrg = null;

            //DateTime dJoinDate = new DateTime();
            //DateTime dRetireDate = new DateTime();
            ////int nRetireDateCnt = 0;

            //if (igrCERTIFICATE.RowCount > 0)
            //{
            //    bCertTypeID = igrCERTIFICATE.GetCellValue("CERT_TYPE_ID");     //증명서 ID
            //    bCertTypeName = igrCERTIFICATE.GetCellValue("CERT_TYPE_NAME"); //증명서
            //    bName = igrCERTIFICATE.GetCellValue("NAME");                   //성명
            //    bPersonID = igrCERTIFICATE.GetCellValue("PERSON_ID");          //사원ID
            //    bDescription = igrCERTIFICATE.GetCellValue("DESCRIPTION");     //용도
            //    bSendOrg = igrCERTIFICATE.GetCellValue("SEND_ORG");            //제출처

            //    dJoinDate = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("JOIN_DATE"));
            //    dRetireDate = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("RETIRE_DATE"));

            //    dPrint_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("PRINT_DATE"));
            //    sPrint_num = igrCERTIFICATE.GetCellValue("PRINT_NUM").ToString();
            //}
            //gridClickPrinting(dPrint_Date, sPrint_num, bCertTypeID, bCertTypeName, bName, bPersonID, dJoinDate, dRetireDate, bDescription, bSendOrg);
        }


        #region ----- Assembly Run Methods ----

        private void AssmblyRun_Manual(object pAssembly_ID, object pCorp_ID, object pPrint_Num, object pPrint_Date)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string vCurrAssemblyFileVersion = string.Empty;

            // Form pMainForm, ISAppInterface pAppInterface, object pCorp_ID, string pUser_Print, string pPrint_Req_Num, object pReq_Date

            //[EAPP_ASSEMBLY_INFO_G.MENU_ENTRY_PROCESS_START]
            IDC_MENU_ENTRY_MANUAL_START.SetCommandParamValue("W_ASSEMBLY_ID", pAssembly_ID);
            IDC_MENU_ENTRY_MANUAL_START.ExecuteNonQuery();

            string vREAD_FLAG = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_READ_FLAG"));
            string vUSER_TYPE = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_USER_TYPE"));
            string vPRINT_FLAG = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_PRINT_FLAG"));

            decimal vASSEMBLY_INFO_ID = iConv.ISDecimaltoZero(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_INFO_ID"));
            string vASSEMBLY_ID = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_ID"));
            string vASSEMBLY_NAME = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_NAME"));
            string vASSEMBLY_FILE_NAME = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_FILE_NAME"));

            string vASSEMBLY_VERSION = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_ASSEMBLY_VERSION"));
            string vDIR_FULL_PATH = iConv.ISNull(IDC_MENU_ENTRY_MANUAL_START.GetCommandParamValue("O_DIR_FULL_PATH"));

            System.IO.FileInfo vFile = new System.IO.FileInfo(vASSEMBLY_FILE_NAME);
            if (vFile.Exists)
            {
                vCurrAssemblyFileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(vASSEMBLY_FILE_NAME).FileVersion;
            }

            vREAD_FLAG = "Y";  //무조건 인쇄

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
                        vFileTransferAdv.UseBinary = true;
                        vFileTransferAdv.KeepAlive = false;
                        if (isAppInterfaceAdv1.AppInterface.AppHostInfo.Passive != "Y")
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

                        //report update//
                        ReportUpdate(vASSEMBLY_INFO_ID);
                    }

                    try
                    {
                        System.Reflection.Assembly vAssembly = System.Reflection.Assembly.LoadFrom(vASSEMBLY_FILE_NAME);
                        Type vType = vAssembly.GetType(vAssembly.GetName().Name + "." + vAssembly.GetName().Name);

                        if (vType != null)
                        {
                            if (vFile.Exists)
                            {
                                vCurrAssemblyFileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(vASSEMBLY_FILE_NAME).FileVersion;
                            }

                            object[] vParam = new object[5];
                            vParam[0] = this.MdiParent;
                            vParam[1] = isAppInterfaceAdv1.AppInterface;
                            vParam[2] = pCorp_ID;       //업체id 
                            vParam[3] = pPrint_Num;     //발급번호
                            vParam[4] = pPrint_Date;    //발급일자

                            object vCreateInstance = Activator.CreateInstance(vType, vParam);
                            Office2007Form vForm = vCreateInstance as Office2007Form;
                            Point vPoint = new Point(30, 30);
                            vForm.Location = vPoint;
                            vForm.StartPosition = FormStartPosition.Manual;
                            vForm.Text = string.Format("{0}[{1}] - {2}", vASSEMBLY_NAME, vASSEMBLY_ID, vCurrAssemblyFileVersion);

                            vForm.Show();
                        }
                        else
                        {
                            MessageBoxAdv.Show("Form Namespace Error");
                        }
                    }
                    catch
                    {
                        //
                    }
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
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
                        if (iConv.ISNull(vRow["REPORT_FILE_NAME"]) != string.Empty)
                        {
                            vReportFileName = iConv.ISNull(vRow["REPORT_FILE_NAME"]);
                            vReportFileNameTarget = string.Format("_{0}", vReportFileName);
                        }
                        if (iConv.ISNull(vRow["REPORT_PATH_FTP"]) != string.Empty)
                        {
                            vPathReportFTP = iConv.ISNull(vRow["REPORT_PATH_FTP"]);
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


    }
}