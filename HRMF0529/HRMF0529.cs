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

namespace HRMF0529
{
    public partial class HRMF0529 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0529()
        {
            InitializeComponent();
        }

        public HRMF0529(Form pMainForm, ISAppInterface pAppInterface)
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
                W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            } 

            if (TB_MAIN.SelectedTab.TabIndex == TP_SUMMARY.TabIndex)
            {
                IDA_4INSURANCE_HEADER.Fill();
                IGR_4INSURANCE_HEADER.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_INTERFACE.TabIndex)
            {
                string vINSUR_TYPE = iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue("INSUR_TYPE"));

                IDA_4INSURANCE_SLIP.Fill();
                IDA_4INSURANCE_SLIP_SUM.Fill();
                 
                int vIDX_INSUR_TYPE = IGR_4INSURANCE_SLIP.GetColumnToIndex("INSUR_TYPE");
                for (int r = 0; r < IGR_4INSURANCE_SLIP.RowCount; r++)
                {
                    if (iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue(r, vIDX_INSUR_TYPE)) == vINSUR_TYPE)
                    {
                        IGR_4INSURANCE_SLIP.CurrentCellMoveTo(r, vIDX_INSUR_TYPE);
                        IGR_4INSURANCE_SLIP.Focus();
                        return;            
                    }
                }
                IGR_4INSURANCE_SLIP.Focus();
            }
        }
          
        private void Set_Dist_Insur_Amount()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            if (iConv.ISNull(V_DIST_STD_PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_DIST_STD_PAY_YYYYMM.Focus();
                return;
            }

            string vINSUR_TYPE = iConv.ISNull(IGR_4INSURANCE_HEADER.GetCellValue("INSUR_TYPE"));
            if (vINSUR_TYPE == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_INSUR_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            //고지금액 반영//
            IDA_4INSURANCE_HEADER.Update();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_4INSURANCE_DIST.SetCommandParamValue("P_INSUR_TYPE", vINSUR_TYPE);
            IDC_SET_4INSURANCE_DIST.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_4INSURANCE_DIST.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_4INSURANCE_DIST.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_4INSURANCE_DIST.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //requery
            Search_DB();
        }

        private void Summary_Amount()
        {
            int vIDX_DR_Amount = IGR_SALARY_SLIP_SUM.GetColumnToIndex("DR_AMOUNT");
            int vIDX_CR_Amount = IGR_SALARY_SLIP_SUM.GetColumnToIndex("CR_AMOUNT");

            decimal vDR_Amount = 0;
            decimal vCR_Amount = 0;
            decimal vGap_Amount = 0;

            for (int vRow = 0; vRow < IGR_SALARY_SLIP_SUM.RowCount; vRow++)
            {
                vDR_Amount = vDR_Amount + iConv.ISDecimaltoZero(IGR_SALARY_SLIP_SUM.GetCellValue(vRow, vIDX_DR_Amount));
                vCR_Amount = vCR_Amount + iConv.ISDecimaltoZero(IGR_SALARY_SLIP_SUM.GetCellValue(vRow, vIDX_CR_Amount));
            }
            vGap_Amount = Math.Abs(vDR_Amount - vCR_Amount) * -1;
            T_DR_SUM.EditValue = vDR_Amount;
            T_CR_SUM.EditValue = vCR_Amount;
            T_GAP_SUM.EditValue = vGap_Amount;
        }

        private void Set_4Insur_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }

            string vPERIOD_NAME= iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue("PERIOD_NAME"));
            if (vPERIOD_NAME == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            string vINSUR_TYPE = iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue("INSUR_TYPE"));
            if (vINSUR_TYPE == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_INSUR_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty; 

            IDC_SET_4INSURANCE_SLIP.SetCommandParamValue("P_PERIOD_NAME", vPERIOD_NAME);
            IDC_SET_4INSURANCE_SLIP.SetCommandParamValue("P_INSUR_TYPE", vINSUR_TYPE);
            IDC_SET_4INSURANCE_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_4INSURANCE_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_4INSURANCE_SLIP.GetCommandParamValue("O_MESSAGE"));
            
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_4INSURANCE_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }     
       
            //requery
            Search_DB();
        }

        private void Cancel_4Insur_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            string vPERIOD_NAME = iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue("PERIOD_NAME"));
            if (vPERIOD_NAME == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            string vINSUR_TYPE = iConv.ISNull(IGR_4INSURANCE_SLIP.GetCellValue("INSUR_TYPE"));
            if (vINSUR_TYPE == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_INSUR_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_4INSURANCE_SLIP.SetCommandParamValue("P_PERIOD_NAME", vPERIOD_NAME);
            IDC_CANCEL_4INSURANCE_SLIP.SetCommandParamValue("P_INSUR_TYPE", vINSUR_TYPE); 
            IDC_CANCEL_4INSURANCE_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_CANCEL_4INSURANCE_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_CANCEL_4INSURANCE_SLIP.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_4INSURANCE_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //requery
            Search_DB();
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

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
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
                    if (IDA_4INSURANCE_HEADER.IsFocused)
                    {
                        IDA_4INSURANCE_HEADER.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_4INSURANCE_HEADER.IsFocused)
                    {
                        IDA_4INSURANCE_HEADER.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_4INSURANCE_HEADER.IsFocused)
                    {
                        IDA_4INSURANCE_HEADER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_4INSURANCE_HEADER.IsFocused)
                    {
                        IDA_4INSURANCE_HEADER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_4INSURANCE_HEADER.IsFocused)
                    {
                        IDA_4INSURANCE_HEADER.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0529_Shown(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            V_DIST_STD_PAY_YYYYMM.EditValue = W_PERIOD_NAME.EditValue;

            DefaultCorporation();              //Default Corp.
        }

        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {          
            Set_4Insur_Slip_Interface();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Cancel_4Insur_Slip_Interface();
        }

        private void BTN_SET_DIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Dist_Insur_Amount();
        }

        #endregion

        #region ----- Lookup Event ------

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ILA_INSUR_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "4INSUR_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }


        private void ilaYYYYMM_0_SelectedRowData(object pSender)
        {
            V_DIST_STD_PAY_YYYYMM.EditValue = W_PERIOD_NAME.EditValue;
        }


        #endregion

        #region ----- Adapter Event -----


        #endregion
    }
}