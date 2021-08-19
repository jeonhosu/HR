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

namespace HRMF0522
{
    public partial class HRMF0522 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0522()
        {
            InitializeComponent();
        }

        public HRMF0522(Form pMainForm, ISAppInterface pAppInterface)
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
            if (iString.ISNull(iedCORP_ID_0.EditValue) == string.Empty)
            {// 업체 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(iedPAY_YYYYMM.EditValue) == string.Empty)
            {// 지급일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedPAY_YYYYMM.Focus();
                return;
            }            
            if (iString.ISNull(iedWAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 선택 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedWAGE_TYPE_NAME.Focus();
                return;
            }
            
            // 그리드 부분 업데이트 처리
            idaMONTH_PAYMENT.OraSelectData.AcceptChanges();
            idaMONTH_PAYMENT.Refillable = true;

            idaMONTH_PAYMENT.Fill();
            //idaALLOWANCE_INFO.Fill();
            //idaDEDUCTION_INFO.Fill();
            //idaDUTY_INFO.Fill();
        }

        #endregion;

        // 인쇄 부분
        // Print 관련 소스 코드 2011.1.15(토)
        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.OraSelectData.Rows.Count;
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
                xlPrinting.OpenFileNameExcel = "HRMF0522_001.xls";
                xlPrinting.XLFileOpen();

                //xlPrinting.PreView();
                
                int vTerritory1 = GetTerritory(igrMONTH_PAYMENT.TerritoryLanguage);
                int vTerritory2 = GetTerritory(igrPAY_ALLOWANCE.TerritoryLanguage);
                int vTerritory3 = GetTerritory(igrPAY_DEDUCTION.TerritoryLanguage);
                int vTerritory4 = GetTerritory(igrDUTY_INFO.TerritoryLanguage);

                string vPeriodFrom = iedPAY_YYYYMM.DateTimeValue.ToString("yyyy-MM-dd", null);
                //string vPeriodTo = END_DATE_0.DateTimeValue.ToString("yyyy-MM-dd", null);

                string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);

                int viCutStart = this.Text.LastIndexOf("]") + 1;
                string vCaption = this.Text.Substring(0, viCutStart);
                //int vPageNumber = xlPrinting.XLWirte(pGrid1, vTerritory1, vPeriodFrom, vUserName, vCaption);  
                int vPageNumber=0;

                int vIndexCheckBox = igrMONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN"); // select의 컬럼 인덱스
                int vTotalRow = igrMONTH_PAYMENT.RowCount; //Grid1의 총 행수

                string vWAGE_TYPE;
                string vPAY_TYPE;

                for (int nRow = 0; nRow < vTotalRow; nRow++)
                {
                    if ((string)igrMONTH_PAYMENT.GetCellValue(nRow, vIndexCheckBox) == "Y")
                    {
                        igrMONTH_PAYMENT.CurrentCellMoveTo(nRow, 0);
                        igrMONTH_PAYMENT.Focus();
                        igrMONTH_PAYMENT.CurrentCellActivate(nRow, 0);

                        vWAGE_TYPE = iString.ISNull(igrMONTH_PAYMENT.GetCellValue(nRow, igrMONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE")));
                        vPAY_TYPE = iString.ISNull(igrMONTH_PAYMENT.GetCellValue(nRow, igrMONTH_PAYMENT.GetColumnToIndex("PAY_TYPE")));

                        if (vWAGE_TYPE == "P1".ToString() && (vPAY_TYPE == "2".ToString() || vPAY_TYPE == "4".ToString()))
                        {
                            xlPrinting.XLWirte(igrMONTH_PAYMENT, nRow, vTerritory1, vPeriodFrom, vUserName, vCaption, 1);
                            xlPrinting.XLWirte(igrPAY_ALLOWANCE, nRow, vTerritory2, vPeriodFrom, vUserName, vCaption, 2);
                            xlPrinting.XLWirte(igrPAY_DEDUCTION, nRow, vTerritory3, vPeriodFrom, vUserName, vCaption, 3);
                            xlPrinting.XLWirte(igrDUTY_INFO, nRow, vTerritory4, vPeriodFrom, vUserName, vCaption, 4);
                            xlPrinting.XLWirte(igrBONUS_ALLOWANCE, nRow, vTerritory2, vPeriodFrom, vUserName, vCaption, 5);
                            xlPrinting.XLWirte(igrBONUS_DEDUCTION, nRow, vTerritory3, vPeriodFrom, vUserName, vCaption, 6);
                        }
                        else
                        {
                            xlPrinting.XLWirte2(igrMONTH_PAYMENT, nRow, vTerritory1, vPeriodFrom, vUserName, vCaption, 1);
                            xlPrinting.XLWirte2(igrPAY_ALLOWANCE, nRow, vTerritory2, vPeriodFrom, vUserName, vCaption, 2);
                            xlPrinting.XLWirte2(igrPAY_DEDUCTION, nRow, vTerritory3, vPeriodFrom, vUserName, vCaption, 3);
                            xlPrinting.XLWirte2(igrDUTY_INFO, nRow, vTerritory4, vPeriodFrom, vUserName, vCaption, 4);
                        }
                        vPageNumber++;
                    }                    
                }
                //xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
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
                    XLPrinting1(); // 출력 함수 호출

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                    // 인쇄 완료 메시지 출력
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0522_Load(object sender, EventArgs e)
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            iedCORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            iedCORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            iedPAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            iedSTART_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            iedEND_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);

            // 그리드 부분 업데이트 처리 위함.
            idaMONTH_PAYMENT.FillSchema();
        }

        // 전체선택 버튼
        private void btnSELECT_ALL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrMONTH_PAYMENT.RowCount; i++)
            {
                igrMONTH_PAYMENT.SetCellValue(i, igrMONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN"), "Y");
            }            
            idaMONTH_PAYMENT.OraSelectData.AcceptChanges();
            idaMONTH_PAYMENT.Refillable = true;
        }

        // 취소 버튼
        private void btnCONFIRM_CANCEL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrMONTH_PAYMENT.RowCount; i++)
            {
                igrMONTH_PAYMENT.SetCellValue(i, igrMONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN"), "N");
            }            
            idaMONTH_PAYMENT.OraSelectData.AcceptChanges();
            idaMONTH_PAYMENT.Refillable = true;
        }

        #endregion

        #region ----- Lookup Event ----- 

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
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

        #endregion

        #region ----- Adapter Event -----
        #endregion
        

    }
}