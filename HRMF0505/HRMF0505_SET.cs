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

namespace HRMF0505
{
    public partial class HRMF0505_SET : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mPROCESS_TYPE = null;
        object mCORP_ID = null;
        object mCORP_NAME = null;
        object mPAY_YYYYMM = null;
        object mWAGE_TYPE = null;
        object mWAGE_TYPE_NAME = null;
        object mDept_ID = null;
        object mDept_Code = null;
        object mDept_Name = null;
        object mFloor_ID = null;
        object mFloor_Name = null;
        object mPerson_ID = null;
        object mPerson_Num = null;
        object mName = null;
        object mCorp_Type = null;

        #endregion;

        #region ----- Constructor -----

        public HRMF0505_SET(ISAppInterface pAppInterface, string pPROCESS_TYPE
                                    , object pCorp_ID, object pCorp_NAME
                                    , object pPay_YYYYMM, object pWage_Type, object pWage_Type_NAME
                                    , object pDept_ID, object pDept_Code, object pDept_Name
                                    , object pFloor_ID, object pFloor_Name
                                    , object pPerson_id, object pPerson_Num, object pName, object pCorp_Type
                                    , object pStd_Date, object pExchange_Rate, object pCurrency_Code)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            

            mPROCESS_TYPE = pPROCESS_TYPE;
            mCORP_ID = pCorp_ID;
            mCORP_NAME = pCorp_NAME;
            mPAY_YYYYMM = pPay_YYYYMM;
            mWAGE_TYPE = pWage_Type;
            mWAGE_TYPE_NAME = pWage_Type_NAME;
            mDept_ID = pDept_ID;
            mDept_Code = pDept_Code;
            mDept_Name = pDept_Name;
            mFloor_ID = pFloor_ID;
            mFloor_Name = pFloor_Name;
            mPerson_ID = pPerson_id;
            mPerson_Num = pPerson_Num;
            mName = pName;
            mCorp_Type = pCorp_Type;

            EXCH_DATE.EditValue = pStd_Date;
            EXCHANGE_RATE.EditValue = pExchange_Rate;
            CURRENCY_CODE.EditValue = pCurrency_Code;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Init_Process_Status()
        {
            if (mPROCESS_TYPE == "CAL")
            {
                TITLE.PromptTextElement[0].TL1_KR = "급/상여 계산";
                TITLE.PromptTextElement[0].Default = "Salary/Bonus Calculate";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[계산]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Calculate]";

                STANDARD_DATE.Visible = true;
                SUPPLY_DATE.Visible = true;
                EXCEPT_YN.Visible = false;
            }
            else if (mPROCESS_TYPE == "CLOSE")
            {
                TITLE.PromptTextElement[0].TL1_KR = "급/상여 마감";
                TITLE.PromptTextElement[0].Default = "Salary/Bonus Close";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[마감]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Close]";

                STANDARD_DATE.Visible = false;
                SUPPLY_DATE.Visible = false;
                EXCEPT_YN.Visible = false;
            }
            else if (mPROCESS_TYPE == "CLOSED_CANCEL")
            {
                TITLE.PromptTextElement[0].TL1_KR = "급/상여 마감 취소";
                TITLE.PromptTextElement[0].Default = "Salary/Bonus Closed Cancel";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[마감 취소]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Closed Cancel]";

                STANDARD_DATE.Visible = false;
                SUPPLY_DATE.Visible = false;
                EXCEPT_YN.Visible = true;
            }
            EXCEPT_YN.Invalidate();
            TITLE.Invalidate();
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
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0505_SET_PAYMENT_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0505_SET_PAYMENT_Shown(object sender, EventArgs e)
        {
            Init_Process_Status();

            CORP_NAME.EditValue = mCORP_NAME;
            CORP_ID.EditValue = mCORP_ID;
            PAY_YYYYMM.EditValue = mPAY_YYYYMM;
            START_DATE.EditValue = iDate.ISMonth_1st(mPAY_YYYYMM.ToString());
            END_DATE.EditValue = iDate.ISMonth_Last(mPAY_YYYYMM.ToString());
            WAGE_TYPE.EditValue = mWAGE_TYPE;
            WAGE_TYPE_NAME.EditValue = mWAGE_TYPE_NAME;
            DEPT_ID.EditValue = mDept_ID;
            DEPT_CODE.EditValue = mDept_Code;
            DEPT_NAME.EditValue = mDept_Name;
            FLOOR_ID.EditValue = mFloor_ID;
            FLOOR_NAME.EditValue = mFloor_Name;
            PERSON_ID.EditValue = mPerson_ID;
            PERSON_NUM.EditValue = mPerson_Num;
            NAME.EditValue = mName;
            CORP_TYPE.EditValue = mCorp_Type;

            idcSUPPLY_DATE.ExecuteNonQuery();
            SUPPLY_DATE.EditValue = idcSUPPLY_DATE.GetCommandParamValue("O_SUPPLY_DATE");
            idcSTANDARD_DATE.ExecuteNonQuery();
            STANDARD_DATE.EditValue = idcSTANDARD_DATE.GetCommandParamValue("O_STANDARD_DATE");

            EXCEPT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void ibtCREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME.Focus();
                return;
            }
            
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            if (mPROCESS_TYPE == "CAL")
            {
                if (iString.ISNull(STANDARD_DATE.EditValue) == string.Empty)
                {// 기준일자 구분
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10110"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    STANDARD_DATE.Focus();
                    return;
                }
                if (iString.ISNull(SUPPLY_DATE.EditValue) == string.Empty)
                {// 지급일자 구분
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10111"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SUPPLY_DATE.Focus();
                    return;
                }

                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    return;
                }

                idcSET_PAYMENT.ExecuteNonQuery();
                vSTATUS = iString.ISNull(idcSET_PAYMENT.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(idcSET_PAYMENT.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (idcSET_PAYMENT.ExcuteError || vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
            }
            else if (mPROCESS_TYPE == "CLOSE")
            {                
                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10383"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    return;
                }

                IDC_PAYMENT_CLOSED.ExecuteNonQuery();
                vSTATUS = iString.ISNull(IDC_PAYMENT_CLOSED.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(IDC_PAYMENT_CLOSED.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (IDC_PAYMENT_CLOSED.ExcuteError || vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
            }
            else if (mPROCESS_TYPE == "CLOSED_CANCEL")
            {
                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10384"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    return;
                }

                IDC_PAYMENT_CLOSED_CANCEL.ExecuteNonQuery();
                vSTATUS = iString.ISNull(IDC_PAYMENT_CLOSED_CANCEL.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(IDC_PAYMENT_CLOSED_CANCEL.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (IDC_PAYMENT_CLOSED_CANCEL.ExcuteError || vSTATUS == "F")
                {
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
            }
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.No;
            this.Close();
        }

        #endregion              

        #region ----- Lookup Event -----

        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        
        #endregion

    }
}