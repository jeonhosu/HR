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

namespace HRMF0412
{
    public partial class HRMF0412_SET : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mPROCESS_TYPE = null;
        object mCORP_ID = null;
        object mCORP_NAME = null;
        object mInsur_YYYYMM = null;
        object mWAGE_TYPE = null;
        object mWAGE_TYPE_NAME = null;
        object mOPERATING_UNIT_DESC = null;
        object mOPERATING_UNIT_ID = null;
        object mFloor_id = null;
        object mFloor_Name = null;
        object mPerson_ID = null;
        object mPerson_Num = null;
        object mName = null;

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0412_SET(ISAppInterface pAppInterface, string pPROCESS_TYPE
                                    , object pCorp_ID, object pCorp_NAME
                                    , object pPay_YYYYMM
                                    , object pWage_Type, object pWage_Type_NAME
                                    , object pOPERATING_UNIT_DESC, object pOPERATING_UNIT_ID 
                                    , object pFloor_id, object pFloor_Name
                                    , object pPerson_id, object pPerson_Num, object pName)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mPROCESS_TYPE = pPROCESS_TYPE;
            mCORP_ID = pCorp_ID;
            mCORP_NAME = pCorp_NAME;
            mInsur_YYYYMM = pPay_YYYYMM;
            mWAGE_TYPE = pWage_Type;
            mWAGE_TYPE_NAME = pWage_Type_NAME;
            mOPERATING_UNIT_DESC = pOPERATING_UNIT_DESC;
            mOPERATING_UNIT_ID = pOPERATING_UNIT_ID;
            mFloor_id = pFloor_id;
            mFloor_Name = pFloor_Name;
            mPerson_ID = pPerson_id;
            mPerson_Num = pPerson_Num;
            mName = pName;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Init_Process_Status()
        {
            if (mPROCESS_TYPE == "CAL")
            {
                TITLE.PromptTextElement[0].TL1_KR = "보험료 계산";
                TITLE.PromptTextElement[0].Default = "Insurance Calculate";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[계산]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Calculate]";

                EXCEPT_YN.Visible = false;
            }
            else if (mPROCESS_TYPE == "CLOSE")
            {
                TITLE.PromptTextElement[0].TL1_KR = "보험료 마감";
                TITLE.PromptTextElement[0].Default = "Insurance Close";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[마감]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Close]";

                EXCEPT_YN.Visible = false;
            }
            else if (mPROCESS_TYPE == "CLOSED_CANCEL")
            {
                TITLE.PromptTextElement[0].TL1_KR = "보험료 마감 취소";
                TITLE.PromptTextElement[0].Default = "Insurance Closed Cancel";
                PT_PROCESS_TYPE.PromptTextElement[0].TL1_KR = "[마감 취소]";
                PT_PROCESS_TYPE.PromptTextElement[0].Default = "[Closed Cancel]";

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

        private void HRMF0412_SET_PAYMENT_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0412_SET_PAYMENT_Shown(object sender, EventArgs e)
        {
            Init_Process_Status();

            CORP_NAME.EditValue = mCORP_NAME;
            CORP_ID.EditValue = mCORP_ID;
            INSUR_YYYYMM.EditValue = mInsur_YYYYMM;
            START_DATE.EditValue = iDate.ISMonth_1st(mInsur_YYYYMM.ToString());
            END_DATE.EditValue = iDate.ISMonth_Last(mInsur_YYYYMM.ToString());
            WAGE_TYPE.EditValue = mWAGE_TYPE;
            WAGE_TYPE_NAME.EditValue = mWAGE_TYPE_NAME;
            W_OPERATING_UNIT_DESC.EditValue = mOPERATING_UNIT_DESC;
            W_OPERATING_UNIT_ID.EditValue = mOPERATING_UNIT_ID;
            FLOOR_ID_0.EditValue = mFloor_id;
            FLOOR_NAME_0.EditValue = mFloor_Name;
            PERSON_ID.EditValue = mPerson_ID;
            PERSON_NUM.EditValue = mPerson_Num;
            NAME.EditValue = mName;

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
            if (iString.ISNull(INSUR_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                INSUR_YYYYMM.Focus();
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
                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10468"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                idcSET_INSURANCE_AMOUNT.ExecuteNonQuery();
                vSTATUS = iString.ISNull(idcSET_INSURANCE_AMOUNT.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(idcSET_INSURANCE_AMOUNT.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (idcSET_INSURANCE_AMOUNT.ExcuteError || vSTATUS == "F")
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else if (mPROCESS_TYPE == "CLOSE")
            {                
                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10383"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                IDC_CLOSED_INSURANCE.ExecuteNonQuery();
                vSTATUS = iString.ISNull(IDC_CLOSED_INSURANCE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(IDC_CLOSED_INSURANCE.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (IDC_CLOSED_INSURANCE.ExcuteError || vSTATUS == "F")
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else if (mPROCESS_TYPE == "CLOSED_CANCEL")
            {
                DialogResult vdlgResult;
                vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10384"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (vdlgResult == DialogResult.No)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                IDC_CANCEL_CLOSED_INSURANCE.ExecuteNonQuery();
                vSTATUS = iString.ISNull(IDC_CANCEL_CLOSED_INSURANCE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iString.ISNull(IDC_CANCEL_CLOSED_INSURANCE.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (IDC_CANCEL_CLOSED_INSURANCE.ExcuteError || vSTATUS == "F")
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
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

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        
        #endregion

    }
}