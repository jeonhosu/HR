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

namespace HRMF0507
{
    public partial class HRMF0507 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0507()
        {
            InitializeComponent();
        }

        public HRMF0507(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            Application.DoEvents();
            idaMONTH_PAYMENT_SPREAD.Fill();
            igrMONTH_PAYMENT.Focus();
            Application.DoEvents();
            Application.UseWaitCursor = false;

            object vObject1 = idaMONTH_PAYMENT_SPREAD.GetSelectParamValue("W_CORP_ID");
            object vObject2 = idaMONTH_PAYMENT_SPREAD.GetSelectParamValue("W_WAGE_TYPE");
            object vObject3 = idaMONTH_PAYMENT_SPREAD.GetSelectParamValue("W_PAY_YYYYMM");
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
        }

        private void Show_Print()
        {
            System.Windows.Forms.DialogResult vdlrResult;

            //if (CORP_ID_0.EditValue == null)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    CORP_NAME_0.Focus();
            //    return;
            //}
            //if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            //{// 급여년월
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    PAY_YYYYMM_0.Focus();
            //    return;
            //}
            //if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            //{// 급상여 구분
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    WAGE_TYPE_NAME_0.Focus();
            //    return;
            //}

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            HRMF0507_PRINT vSHOW_PRINT = new HRMF0507_PRINT(isAppInterfaceAdv1.AppInterface
                                                           , CORP_ID_0.EditValue
                                                           , CORP_NAME_0.EditValue
                                                           , PAY_YYYYMM_0.EditValue
                                                           , WAGE_TYPE_0.EditValue
                                                           , WAGE_TYPE_NAME_0.EditValue
                                                           , DEPT_NAME_0.EditValue
                                                           , igrMONTH_PAYMENT);
            vdlrResult = vSHOW_PRINT.ShowDialog();
            vSHOW_PRINT.Dispose();


            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

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
                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    string vOpenExcelFileName = "HRMF0507_002.xls";
                    string vSaveExcelFileName = saveFileDialog1.FileName;

                    XLExport mExport = new XLExport();
                    int vTerritory = GetTerritory(pGrid.TerritoryLanguage);
                    bool vbXLSaveOK = mExport.ExcelExport(pGrid, vTerritory, vOpenExcelFileName, vSaveExcelFileName, this.Text, this);
                    if (vbXLSaveOK == true)
                    {
                        vMessage = string.Format("Save OK [{0}]", vSaveExcelFileName);
                        isAppInterfaceAdv1.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                    }
                    else
                    {
                        vMessage = string.Format("Save Err [{0}]", vSaveExcelFileName);
                        isAppInterfaceAdv1.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                    }

                    System.Windows.Forms.Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Windows.Forms.Application.DoEvents();
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
            int vPageNumber = 0;

            int vCountRowGrid = igrMONTH_PAYMENT.RowCount;

            if (vCountRowGrid < 1)
            {
                return;
            }

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0507_001.xls";
                bool IsOpen = xlPrinting.XLFileOpen();
                if (IsOpen == true)
                {
                    isAppInterfaceAdv1.OnAppMessage("Printing Start...");

                    System.Windows.Forms.Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();

                    int vTerritory = GetTerritory(igrMONTH_PAYMENT.TerritoryLanguage);

                    string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);
                    vUserName = isAppInterfaceAdv1.DISPLAY_NAME;
                    int viCutStart = vUserName.LastIndexOf("(");
                    vUserName = vUserName.Substring(0, viCutStart);

                    string vCORP_NAME = CORP_NAME_0.EditValue as string;
                    string vYYYYMM = PAY_YYYYMM_0.EditValue as string;
                    string vWageTypeName = WAGE_TYPE_NAME_0.EditValue as string;
                    string vDepartment_NAME = DEPT_NAME_0.EditValue as string;
                    vPageNumber = xlPrinting.XLWirte(igrMONTH_PAYMENT, vTerritory, vUserName, vCORP_NAME, vYYYYMM, vWageTypeName, vDepartment_NAME);
                    
                    ////[PRINTER]
                    //xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    ////xlPrinting.Printing(3, 4);


                    ////[SAVE]
                    xlPrinting.Save("Out_"); //저장 파일명


                    //[PREVIEW]
                    //xlPrinting.PreView();
                    //-------------------------------------------------------------------------
                }
                else
                {
                    xlPrinting.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }

            xlPrinting.Dispose();

            vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
            isAppInterfaceAdv1.OnAppMessage(vMessageText);

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

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
                    //Show_Print();
                    XLPrinting1();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    ExportXL(igrMONTH_PAYMENT);
                }
            }
        }

        #endregion;
        
        #region ----- Form Event -----
        private void HRMF0507_Load(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);

            DefaultCorporation();              //Default Corp.
        }

        #endregion

        #region ----- Lookup Event -----
        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("PAY_GRADE", "N");
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }
        #endregion
        
    }
}