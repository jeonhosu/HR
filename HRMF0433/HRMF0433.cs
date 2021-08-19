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

namespace HRMF2525
{
    public partial class HRMF2525 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF2525()
        {
            InitializeComponent();
        }

        public HRMF2525(Form pMainForm, ISAppInterface pAppInterface)
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
                ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                idcDEFAULT_CORP.ExecuteNonQuery();

                W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
                W_CORP_NAME.BringToFront();
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
            if (iConv.ISNull(W_PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PAY_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(W_WAGE_TYPE.EditValue) == String.Empty)
            {// 지급구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_LIST.TabIndex)
            {
                IDA_SALARY_SUM.Fill();
                IGR_SALARY_SUM.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PERSON.TabIndex)
            {
                IDA_SALARY_PERSON.Fill();
                IGR_SALARY_PERSON.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_SALARY_SLIP.TabIndex)
            {
                IDA_SALARY_SLIP_SUM_CN.Fill();
                Summary_Amount();
                IGR_SALARY_SLIP_SUM.Focus();
            } 
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

        private void Set_Salary_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PAY_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(W_WAGE_TYPE.EditValue) == String.Empty)
            {// 지급구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_SALARY_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_SALARY_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_SALARY_SLIP.GetCommandParamValue("O_MESSAGE"));
            
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_SALARY_SLIP.ExcuteError || vStatus == "F")
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

        private void Cancel_Salary_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PAY_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(W_WAGE_TYPE.EditValue) == String.Empty)
            {// 지급구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_SALARY_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_CANCEL_SALARY_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_CANCEL_SALARY_SLIP.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_SALARY_SLIP.ExcuteError || vStatus == "F")
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_SALARY_PERSON.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SALARY_SUM.IsFocused)
                    {
                        IDA_SALARY_SUM.Cancel();
                    }
                    else if(IDA_SALARY_PERSON.IsFillCompleted)
                    {
                        IDA_SALARY_PERSON.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF2525_Shown(object sender, EventArgs e)
        {
            W_PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            DefaultCorporation();              //Default Corp.
        }

        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Salary_Slip_Interface();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Cancel_Salary_Slip_Interface();
        }

        #endregion

        #region ----- Lookup Event ------

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DEPT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_COST_CENTER_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_FLOOR_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }

        private void ILA_COST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DIR_INDIR_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("DIR_INDIR_TYPE", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_SALARY_PERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if(iConv.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_SALARY_PERSON, IGR_SALARY_PERSON.GetColumnToIndex("MONTH_PAYMENT_ID"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["DIR_INDIR_TYPE"]) == string.Empty)
            {
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_SALARY_PERSON, IGR_SALARY_PERSON.GetColumnToIndex("DIR_INDIR_TYPE"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["COST_CENTER_ID"]) == string.Empty)
            {
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_SALARY_PERSON, IGR_SALARY_PERSON.GetColumnToIndex("COST_CENTER_ID"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        #endregion

    }
}