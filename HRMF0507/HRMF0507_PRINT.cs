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

namespace HRMF0507
{
    public partial class HRMF0507_PRINT : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCORP_ID;
        object mCORP_NAME;
        object mPAY_YYYYMM;
        object mWAGE_TYPE;
        object mWAGE_TYPE_NAME;
        object mDepartment_NAME;

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGrid;

        #endregion;

        #region ----- Constructor -----

        public HRMF0507_PRINT(ISAppInterface pAppInterface
                             , object pCorp_ID
                             , object pCorp_NAME
                             , object pPay_YYYYMM
                             , object pWage_Type
                             , object pWage_Type_NAME
                             , object pDepartment_NAME
                             , InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mCORP_ID = pCorp_ID;
            mCORP_NAME = pCorp_NAME;
            mPAY_YYYYMM = pPay_YYYYMM;
            mWAGE_TYPE = pWage_Type;
            mWAGE_TYPE_NAME = pWage_Type_NAME;
            mDepartment_NAME = pDepartment_NAME;

            mGrid = pGrid;
        }

        #endregion;

        #region ----- Private Methods ----


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
            //string vMessageText = string.Empty;
            //int vPageNumber = 0;

            //int vCountRowGrid = mGrid.RowCount;

            //if (vCountRowGrid < 1)
            //{
            //    return;
            //}

            //XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1);

            //try
            //{
            //    //-------------------------------------------------------------------------
            //    xlPrinting.OpenFileNameExcel = "HRMF0507_001.xls";
            //    bool IsOpen = xlPrinting.XLFileOpen();
            //    if (IsOpen == true)
            //    {
            //        isAppInterfaceAdv1.OnAppMessage("Printing Start...");

            //        System.Windows.Forms.Application.UseWaitCursor = true;
            //        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //        System.Windows.Forms.Application.DoEvents();

            //        int vTerritory = GetTerritory(mGrid.TerritoryLanguage);

            //        string vUserName = string.Format("[{0}]{1}", isAppInterfaceAdv1.DEPT_NAME, isAppInterfaceAdv1.DISPLAY_NAME);
            //        vUserName = isAppInterfaceAdv1.DISPLAY_NAME;
            //        int viCutStart = vUserName.LastIndexOf("(");
            //        vUserName = vUserName.Substring(0, viCutStart);

            //        string vCORP_NAME = mCORP_NAME as string;
            //        string vYYYYMM = mPAY_YYYYMM as string;
            //        string vWageTypeName = mWAGE_TYPE_NAME as string;
            //        string vDepartment_NAME = mDepartment_NAME as string;
            //        vPageNumber = xlPrinting.XLWirte(mGrid, vTerritory, vUserName, vCORP_NAME, vYYYYMM, vWageTypeName, vDepartment_NAME);

            //        ////[PRINTER]
            //        //xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
            //        ////xlPrinting.Printing(3, 4);


            //        ////[SAVE]
            //        xlPrinting.Save("Out_"); //저장 파일명


            //        //[PREVIEW]
            //        //xlPrinting.PreView();
            //        //-------------------------------------------------------------------------
            //    }
            //    else
            //    {
            //        xlPrinting.Dispose();
            //    }
            //}
            //catch (System.Exception ex)
            //{
            //    string vMessage = ex.Message;
            //    xlPrinting.Dispose();
            //}

            //xlPrinting.Dispose();

            //vMessageText = string.Format("Print End! [Page : {0}]", vPageNumber);
            //isAppInterfaceAdv1.OnAppMessage(vMessageText);

            //System.Windows.Forms.Application.UseWaitCursor = false;
            //this.Cursor = System.Windows.Forms.Cursors.Default;
            //System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

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

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0507_PRINT_Load(object sender, EventArgs e)
        {
            CORP_NAME.EditValue = mCORP_NAME;
            CORP_ID.EditValue = mCORP_ID;
            PAY_YYYYMM.EditValue = mPAY_YYYYMM;
            WAGE_TYPE.EditValue = mWAGE_TYPE;
            WAGE_TYPE_NAME.EditValue = mWAGE_TYPE_NAME;
        }


        private void ibtPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (CORP_ID.EditValue == null)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    CORP_NAME.Focus();
            //    return;
            //}
            //if (iString.ISNull(PAY_YYYYMM.EditValue) == String.Empty)
            //{// 급여년월
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    PAY_YYYYMM.Focus();
            //    return;
            //}
            //if (iString.ISNull(WAGE_TYPE.EditValue) == string.Empty)
            //{// 급상여 구분
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    WAGE_TYPE_NAME.Focus();
            //    return;
            //}

            XLPrinting1();
        }

        private void ibtCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion              

        #region ----- Lookup Event -----
        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        #endregion
        
        
    }
}