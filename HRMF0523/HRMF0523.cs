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

namespace HRMF0523
{
    public partial class HRMF0523 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0523()
        {
            InitializeComponent();
        }

        public HRMF0523(Form pMainForm, ISAppInterface pAppInterface)
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
            if (iString.ISNull(iedWAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 선택 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedWAGE_TYPE_NAME.Focus();
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

            if (iString.ISNull(W_REPRE_NUM.EditValue) == string.Empty)
            {// 급상여 선택 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10027"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_REPRE_NUM.Focus();
                return;
            }

            


            // 그리드 부분 업데이트 처리
            IDA_MONTH_PAYMENT.OraSelectData.AcceptChanges();
            IDA_MONTH_PAYMENT.Refillable = true;

            IDA_MONTH_PAYMENT.Fill();
            IDA_PAY_ALLOWANCE.Fill();
            idaDUTY_INFO.Fill();

        }

        #endregion;

        // 인쇄 부분
        // Print 관련 소스 코드 2011.1.15(토)
        // Print 관련 소스 코드 2011.5.11(수) 수정
        #region ----- Convert String Method ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pCourse)
        {
            System.DateTime vStartTime = DateTime.Now;

            string vMessageText = string.Empty;

            string vBoxCheck = string.Empty;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;

            int vCountCheck = 0;

            object vObject = null;

            int vCountRow = igrMONTH_PAYMENT.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            int vIndexWAGE_TYPE = igrMONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = igrMONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");


            //-------------------------------------------------------------------------------------

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            igbCONDITION.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0523_002.xlsx";
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(pCourse, igrMONTH_PAYMENT, IDA_PAY_ALLOWANCE, idaPAY_DEDUCTION, idaMONTH_DUTY, idaMONTH_OT);
            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessageText = string.Format("Printing End [Total Page : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            this.Cursor = System.Windows.Forms.Cursors.Default;
            igbCONDITION.Cursor = System.Windows.Forms.Cursors.Default;
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
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_MONTH_PAYMENT.IsFocused)
                    {
                        IDA_MONTH_PAYMENT.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_MONTH_PAYMENT.IsFocused)
                    {
                        IDA_MONTH_PAYMENT.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_MONTH_PAYMENT.IsFocused)
                    {
                        IDA_MONTH_PAYMENT.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_MONTH_PAYMENT.IsFocused)
                    {
                        IDA_MONTH_PAYMENT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_MONTH_PAYMENT.IsFocused)
                    {
                        IDA_MONTH_PAYMENT.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT"); // 출력 함수 호출
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                  /// XLPrinting_1("FILE"); // 출력 함수 호출
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0523_Load(object sender, EventArgs e)
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            iedCORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            iedCORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            idcDEFAULT_WORK_TYPE.SetCommandParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            idcDEFAULT_WORK_TYPE.ExecuteNonQuery();
            iedWAGE_TYPE_NAME.EditValue = idcDEFAULT_WORK_TYPE.GetCommandParamValue("O_CODE_NAME");
            iedWAGE_TYPE.EditValue = idcDEFAULT_WORK_TYPE.GetCommandParamValue("O_CODE");

            if (DateTime.Today.Day < 15)
            {
                iedPAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today, -1, 0);
                iedSTART_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today, -1, 0).ToString("yyyy-MM-dd");
                iedEND_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today, -1, 0).ToString("yyyy-MM-dd");
            }
            else
            {
                iedPAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today, 0);
                iedSTART_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today, 0).ToString("yyyy-MM-dd");
                iedEND_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today, 0).ToString("yyyy-MM-dd");
            } 
            iedPERSON_ID_0.EditValue = isAppInterfaceAdv1.AppInterface.PersonId;
            string t = iedPERSON_ID_0.EditValue.ToString();

            iedNAME_0.EditValue = isAppInterfaceAdv1.AppInterface.DisplayName;


            IDA_MONTH_PAYMENT.FillSchema();
            IDA_PAY_ALLOWANCE.FillSchema();
            idaDUTY_INFO.FillSchema();

            isAppInterfaceAdv1.OnAppMessage("");
        }

        #endregion;

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (DateTime.Today.Day < 15)
            {
                ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", iDate.ISDate_Month_Add(DateTime.Today, -3).ToString("yyyy-MM-dd"));
            }
            else
            {
                ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", iDate.ISDate_Month_Add(DateTime.Today, -4).ToString("yyyy-MM-dd"));
            }
            
            if (DateTime.Today.Day < 15)
            {
                ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISDate_Month_Add(DateTime.Today, -1).ToString("yyyy-MM-dd"));
            }
            else
            {
                ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISGetDate(DateTime.Today).ToString("yyyy-MM-dd"));
            }
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
    }
}