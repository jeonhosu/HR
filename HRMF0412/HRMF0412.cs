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
    public partial class HRMF0412 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #region ----- Constructor -----
        public HRMF0412(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Property / Method ----

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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (INSUR_YYYYMM_0.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            string vPERSON_ID = iString.ISNull(IGR_INSUR_AMOUNT.GetCellValue("PERSON_ID"));
            IDA_INSUR_AMOUNT.Fill();

            //합계 표시 
            Insurance_Summary();

            IGR_INSUR_AMOUNT.Focus();
            if (IGR_INSUR_AMOUNT.RowCount > 1)
            {
                int vIDX_PERSON_ID = IGR_INSUR_AMOUNT.GetColumnToIndex("PERSON_ID");
                for (int i = 0; i < IGR_INSUR_AMOUNT.RowCount; i++)
                {
                    if (vPERSON_ID == iString.ISNull(IGR_INSUR_AMOUNT.GetCellValue(i, vIDX_PERSON_ID)))
                    {
                        IGR_INSUR_AMOUNT.CurrentCellMoveTo(i, IGR_INSUR_AMOUNT.GetColumnToIndex("NAME"));
                        return;
                    }
                }
            }
        }

        private void Insurance_Summary()
        {
            decimal mPerson_Social_Amount = 0;
            decimal mCompany_Social_Amount = 0;
            decimal mPerson_Unemployed_Amount = 0;
            decimal mCompany_Unemployed_Amount = 0;
            decimal mPerson_Medical_Amount = 0;
            decimal mCompany_Medical_Amount = 0;

            //Index.
            int mIDX_Person_Social = IGR_INSUR_AMOUNT.GetColumnToIndex("P_SOCIAL_INSUR");
            int mIDX_Company_Social = IGR_INSUR_AMOUNT.GetColumnToIndex("C_SOCIAL_INSUR");
            int mIDX_Person_Unemployed = IGR_INSUR_AMOUNT.GetColumnToIndex("P_UNEMPLOYED_INSUR");
            int mIDX_Company_Unemployed = IGR_INSUR_AMOUNT.GetColumnToIndex("C_UNEMPLOYED_INSUR");
            int mIDX_Person_Medical = IGR_INSUR_AMOUNT.GetColumnToIndex("P_MEDICAL_INSUR");
            int mIDX_Company_Medical = IGR_INSUR_AMOUNT.GetColumnToIndex("C_MEDICAL_INSUR");
            for (int i = 0; i < IGR_INSUR_AMOUNT.RowCount; i++)
            {
                //SOCIAL
                mPerson_Social_Amount = iString.ISDecimaltoZero(mPerson_Social_Amount) +
                                       iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Person_Social));
                mCompany_Social_Amount = iString.ISDecimaltoZero(mCompany_Social_Amount) +
                                        iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Company_Social));

                //UNEMPLOYED
                mPerson_Unemployed_Amount = iString.ISDecimaltoZero(mPerson_Unemployed_Amount) +
                                        iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Person_Unemployed));
                mCompany_Unemployed_Amount = iString.ISDecimaltoZero(mCompany_Unemployed_Amount) +
                                        iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Company_Unemployed));

                //MEDICAL
                mPerson_Medical_Amount = iString.ISDecimaltoZero(mPerson_Medical_Amount) +
                                        iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Person_Medical));
                mCompany_Medical_Amount = iString.ISDecimaltoZero(mCompany_Medical_Amount) +
                                        iString.ISDecimaltoZero(IGR_INSUR_AMOUNT.GetCellValue(i, mIDX_Company_Medical));

            }
            PERSON_SOCIAL_INSUR.EditValue = mPerson_Social_Amount;
            COMPANY_SOCIAL_INSUR.EditValue = mCompany_Social_Amount;

            PERSONAL_UNEMPLOYED_INSUR.EditValue = mPerson_Unemployed_Amount;
            COMPANY_UNEMPLOYED_INSUR.EditValue = mCompany_Unemployed_Amount;

            PERSONAL_MEDICAL_INSUR.EditValue = mPerson_Medical_Amount;
            COMPANY_MEDICAL_INSUR.EditValue = mCompany_Medical_Amount;
        }

        #endregion

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

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_INSUR_AMOUNT.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_INSUR_AMOUNT.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_INSUR_AMOUNT.IsFocused)
                    {
                       IDA_INSUR_AMOUNT.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_INSUR_AMOUNT.IsFocused)
                    {
                        IDA_INSUR_AMOUNT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0412_Load(object sender, EventArgs e)
        {
            // FillSchema
            IDA_INSUR_AMOUNT.FillSchema();
        }

        private void HRMF0412_Shown(object sender, EventArgs e)
        {
            INSUR_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]

            DefaultCorporation();                  // Corp Default Value Setting.

            // LEAVE CLOSE TYPE SETTING
            ildAPPROVAL_STATUS_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ildAPPROVAL_STATUS_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            APPROVAL_STATUS_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            APPROVAL_STATUS_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();
        }

        private void SET_INSURANCE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(INSUR_YYYYMM_0.EditValue) == String.Empty)
            {// 보험료년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                INSUR_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0412_SET = new HRMF0412_SET(isAppInterfaceAdv1.AppInterface, "CAL"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , INSUR_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0412_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0412_SET.ShowDialog();
            vHRMF0412_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtSET_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(INSUR_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                INSUR_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0412_SET = new HRMF0412_SET(isAppInterfaceAdv1.AppInterface, "CLOSE"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , INSUR_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0412_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0412_SET.ShowDialog();
            vHRMF0412_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtSET_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(INSUR_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                INSUR_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0412_SET = new HRMF0412_SET(isAppInterfaceAdv1.AppInterface, "CLOSED_CANCEL"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , INSUR_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0412_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0412_SET.ShowDialog();
            vHRMF0412_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void IDA_INSUR_AMOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["INSUR_YYYYMM"]) == string.Empty)
            {//
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(INSUR_YYYYMM_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {// cc
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WAGE_TYPE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_INSUR_AMOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void IDA_INSUR_AMOUNT_UpdateCompleted(object pSender)
        {
            //합계 표시 
            Insurance_Summary();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion
        
    }
}