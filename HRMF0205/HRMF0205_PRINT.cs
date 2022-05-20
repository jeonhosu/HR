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
    public partial class HRMF0205_PRINT : Office2007Form
    {
        ISHR.isCertificatePrint mPrintInfo;
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private string mREPORT_TYPE = string.Empty;
        private string mREPORT_FILENAME = string.Empty;

        public HRMF0205_PRINT(Form pMainForm, ISHR.isCertificatePrint pPrintInfo, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mPrintInfo = new ISHR.isCertificatePrint();
            mPrintInfo = pPrintInfo;
            mPrintInfo.ISPrinting += ISOnPrint;


            V_RB_KO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_LANG_CODE.EditValue = V_RB_KO.RadioCheckedString;

        }

        private void ISOnPrint(string pFormID)
        {
            //iedPRINT_NUM.EditValue = mPrintInfo.Print_Num;
            //iedPRINT_DATE.EditValue = mPrintInfo.Print_Date;
            iedPRINT_DATE.EditValue = DateTime.Today;
            iedPRINT_COUNT.EditValue = Convert.ToInt32(1);
            //if (mPrintInfo.Print_Num != null)
            //{
                iedCERT_TYPE_NAME.EditValue = mPrintInfo.Cert_Type_Name;
                iedCERT_TYPE_ID.EditValue = mPrintInfo.Cert_Type_ID;
                iedNAME.EditValue = mPrintInfo.Name;
                iedPERSON_ID.EditValue = mPrintInfo.Person_ID;
                if (mPrintInfo.Join_Date.Year == 1)
                {
                    iedJOIN_DATE.EditValue = DBNull.Value;
                }
                else
                {
                    iedJOIN_DATE.EditValue = mPrintInfo.Join_Date;
                }                
                if (mPrintInfo.Retire_Date.Year == 1)
                {
                    iedRETIRE_DATE.EditValue = DBNull.Value;
                }
                else
                {
                    iedRETIRE_DATE.EditValue = mPrintInfo.Retire_Date;
                }
                iedDESCRIPTION.EditValue = mPrintInfo.Description;
                iedSEND_ORG.EditValue = mPrintInfo.Send_Org;
                //iedPRINT_COUNT.EditValue = mPrintInfo.Print_Count;
                iedPRINT_COUNT.EditValue = 1;
            //}
        }

        /*
        private void Print_Certificate(object pPrint_num)
        {
            idaCERTIFICATE_INFO.Fill(); // 증명서 인쇄 폼 내에 그리드 부분에 삽입될 데이터 처리.

            XLPrinting1(); // 출력 함수 호출
        }
        */


        // Print 관련 소스 코드 2010.12.21
        /*
        #region ----- XL Export Methods ----
        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.OraSelectData.Rows.Count; // (1)
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xls";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName);
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBoxAdv.Show(vsMessage);
                }
                xlExport.XLClose();
            }
        }

        #endregion;
        */

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

        private void XLPrinting_Main()
        {
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0205");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            mREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            mREPORT_FILENAME = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_FILE_NAME"));

            XLPrinting1();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void XLPrinting1()
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
                    xlPrinting.OpenFileNameExcel = "HRMF0205_001.xlsx";
                }
                xlPrinting.XLFileOpen(); 

                //xlPrinting.PreView();

                int vTerritory = GetTerritory(pGrid.TerritoryLanguage);
                string vPeriodFrom = iedPRINT_DATE.DateTimeValue.ToString("yyyy-MM-dd", null);
                //string vPeriodTo = END_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);

                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);

                int viCutStart = this.Text.LastIndexOf("]") + 1;
                string vCaption = this.Text.Substring(0, viCutStart);
                string vREPRE_FLAG = icb_REPRE_FLAG.CheckBoxString ;

                int nPrintTotalCnt = iString.ISNumtoZero(iedPRINT_COUNT.EditValue);
                xlPrinting.XLWirte(pGrid, pGrid_History, nPrintTotalCnt, vTerritory, vPeriodFrom, vUserName, vCaption, V_LANG_CODE.EditValue.ToString());

                xlPrinting.Printing(1, nPrintTotalCnt); //시작 페이지 번호, 종료 페이지 번호
                //xlPrinting.Printing(3, 4);


                //xlPrinting.Save("Cashier_"); //저장 파일명

                //xlPrinting.PreView();

                xlPrinting.Dispose();
                //-------------------------------------------------------------------------

                //vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
                //isAppInterfaceAdv1.OnAppMessage(vMessageText);
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        #endregion;


        #region ----- Form Event -----
        private void HRMF0205_PRINT_Load(object sender, EventArgs e)
        {
            iedPRINT_DATE.Focus();
        }

        private void HRMF0205_PRINT_FormClosed(object sender, FormClosedEventArgs e)
        {
            mPrintInfo.ISPrintedEvent(mPrintInfo.FormID);
        }

        private void ibtPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 증명서 발급
            if (iedCERT_TYPE_ID.EditValue == null)
            {// 증명서 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10033"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCERT_TYPE_NAME.Focus();
                return;
            }

            if (iedPERSON_ID.EditValue == null)
            {// 사원 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCERT_TYPE_NAME.Focus();
                return;
            }

            if (string.IsNullOrEmpty(iedDESCRIPTION.EditValue.ToString()))
            {// 용도 입력
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10034"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCERT_TYPE_NAME.Focus();
                return;
            }

            // 인쇄 메서드 호출.

            // 인쇄 결과 저장.     
            idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_CORP_ID", mPrintInfo.Corp_ID);
            idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
            idcCERTIFICATE_PRINT_INSERT.ExecuteNonQuery();
            iedPRINT_NUM.EditValue = idcCERTIFICATE_PRINT_INSERT.GetCommandParamValue("P_PRINT_NUM");

            // 인쇄발급 루틴 추가 //
            if (iString.ISNull(iedPRINT_NUM.EditValue) == string.Empty)
            {// 인쇄번호 없음. 인쇄 실패.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //Print_Certificate(iedPRINT_NUM.EditValue); // 증명서 인쇄 폼 안에 있는 그리드 관련 함수
            idaCERTIFICATE_INFO.Fill(); // 증명서 인쇄 폼 내에 그리드 부분에 삽입될 데이터 처리.

            XLPrinting_Main();
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(isMessageAdapter1.ReturnText("FCM_10035"));
            // 인쇄 완료 메시지 출력

            iedPRINT_NUM.EditValue = null;
            iedRETIRE_DATE.EditValue = null;
            iedCERT_TYPE_ID.EditValue = null;
            iedCERT_TYPE_NAME.EditValue = null;
            iedPERSON_ID.EditValue = null;
            PERSON_NUM.EditValue = null;
            iedNAME.EditValue = null;
            iedJOIN_DATE.EditValue = null;
            iedRETIRE_DATE.EditValue = null;
            iedDESCRIPTION.EditValue = null;
            iedSEND_ORG.EditValue = null;
            iedPRINT_COUNT.EditValue = Convert.ToInt32(1);            
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        private void ilaCERT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 10 ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (iedEMPLOYE_TYPE.EditValue.ToString() == "1".ToString())
            {
                ildPERSON.SetLookupParamValue("W_START_DATE", iedPRINT_DATE.EditValue);
                ildPERSON.SetLookupParamValue("W_END_DATE", iedPRINT_DATE.EditValue);
            }
            else
            {
                ildPERSON.SetLookupParamValue("W_START_DATE", DateTime.Parse("2001-01-01"));
                ildPERSON.SetLookupParamValue("W_END_DATE", DateTime.Today);
            }
            ildPERSON.SetLookupParamValue("W_CORP_ID", mPrintInfo.Corp_ID);
        }
        #endregion

        private void V_RB_KO_CheckChanged(object sender, EventArgs e)
        {
            if (V_RB_EN.Checked == true)
            {
                V_LANG_CODE.EditValue = V_RB_EN.RadioCheckedString;
            }
        }
    }
}