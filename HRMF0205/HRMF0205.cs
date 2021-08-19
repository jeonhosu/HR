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
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
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
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void SEARCH_DB()
        {
            IDA_CERTIFICATE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_CERTIFICATE.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_CERTIFICATE.Fill();
        }

        private void isOnPrinting(DateTime pPrint_Date, string pPrint_num)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            if (STD_DATE_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE_0.Focus();
                return;
            }

            ISHR.isCertificatePrint vCertificatePrint = new ISHR.isCertificatePrint();
            vCertificatePrint.FormID = this.Name;
            vCertificatePrint.Corp_ID = Convert.ToInt32( CORP_ID_0.EditValue);
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
            vCertificatePrint.Corp_ID = Convert.ToInt32(CORP_ID_0.EditValue);
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
                    DateTime dPrint_Date = DateTime.Today;
                    string sPrint_num = null;

                    isOnPrinting(dPrint_Date, sPrint_num);
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
            STD_DATE_0.EditValue = DateTime.Today;

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
            ildPERSON.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
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
            DateTime dPrint_Date = DateTime.Today; //발급일자(기준일자)      
            string sPrint_num = null;              //발급번호

            object bCertTypeID = null;
            object bCertTypeName = null;
            object bName = null;
            object bPersonID = null;
            object bDescription = null;
            object bSendOrg = null;

            DateTime dJoinDate = new DateTime();
            DateTime dRetireDate = new DateTime();
            //int nRetireDateCnt = 0;

            if (igrCERTIFICATE.RowCount > 0)
            {
                bCertTypeID = igrCERTIFICATE.GetCellValue("CERT_TYPE_ID");     //증명서 ID
                bCertTypeName = igrCERTIFICATE.GetCellValue("CERT_TYPE_NAME"); //증명서
                bName = igrCERTIFICATE.GetCellValue("NAME");                   //성명
                bPersonID = igrCERTIFICATE.GetCellValue("PERSON_ID");          //사원ID
                bDescription = igrCERTIFICATE.GetCellValue("DESCRIPTION");     //용도
                bSendOrg = igrCERTIFICATE.GetCellValue("SEND_ORG");            //제출처

                dJoinDate = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("JOIN_DATE"));
                dRetireDate = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("RETIRE_DATE"));

                dPrint_Date = Convert.ToDateTime(igrCERTIFICATE.GetCellValue("PRINT_DATE"));
                sPrint_num = igrCERTIFICATE.GetCellValue("PRINT_NUM").ToString();
            }
            gridClickPrinting(dPrint_Date, sPrint_num, bCertTypeID, bCertTypeName, bName, bPersonID, dJoinDate, dRetireDate, bDescription, bSendOrg);
        }
   
    }
}