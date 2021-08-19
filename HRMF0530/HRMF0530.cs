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

namespace HRMF0530
{
    public partial class HRMF0530 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0530()
        {
            InitializeComponent();
        }

        public HRMF0530(Form pMainForm, ISAppInterface pAppInterface)
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
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                idcDEFAULT_CORP.ExecuteNonQuery();
                W_CORP_DESC.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
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
                W_CORP_DESC.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_CHECK.TabIndex)
            {
                IDA_SALARY_CLOSED.Fill();
                IDA_SALARY_ITEM.Fill();
                Search_DB_Detail(string.Empty, string.Empty, string.Empty);
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_ITEM_VALIDATE.TabIndex)
            {
                IDA_SALARY_ITEM_AMOUNT.Fill();
                IDA_SALARY_ITEM_DETAIL.Fill();
                Search_DB_Item_Detail(string.Empty, 0, string.Empty);               
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_TOTAL_VALIDATE.TabIndex)
            {
                IDA_SALARY_TOTAL_AMOUNT.Fill();
                IDA_SALARY_TOTAL_DETAIL.Fill();
            }
        }

        private void Search_DB_Detail(string pDETAIL_TYPE, string pITEM_CODE, string pPROMPT_TEXT )
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_DESC.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            PT_TITLE_0.PromptTextElement[0].Default = string.Format("=>{0}", pPROMPT_TEXT);
            PT_TITLE_0.Refresh();

            IDA_SALARY_DETAIL.SetSelectParamValue("W_DETAIL_TYPE", pDETAIL_TYPE);
            IDA_SALARY_DETAIL.SetSelectParamValue("W_ITEM_CODE", pITEM_CODE);
            IDA_SALARY_DETAIL.Fill();
        }

        private void Search_DB_Item_Detail(string pDETAIL_TYPE, int pITEM_ID, string pPROMPT_TEXT)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_DESC.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            PT_TITLE_1.PromptTextElement[0].Default = string.Format("=>{0}", pPROMPT_TEXT);
            PT_TITLE_1.Refresh();

            IDA_SALARY_ITEM_DETAIL.SetSelectParamValue("W_DETAIL_TYPE", pDETAIL_TYPE);
            IDA_SALARY_ITEM_DETAIL.SetSelectParamValue("W_ITEM_ID", pITEM_ID);
            IDA_SALARY_ITEM_DETAIL.Fill();
        }

        private void Search_DB_Toal_Detail(object pALLOWANCE_TYPE, object pPERSON_ID, string pPROMPT_TEXT)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_DESC.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            PT_TITLE_2.PromptTextElement[0].Default = string.Format("=>{0}", pPROMPT_TEXT);
            PT_TITLE_2.Refresh();

            IDA_SALARY_TOTAL_DETAIL.SetSelectParamValue("W_ALLOWANCE_TYPE", pALLOWANCE_TYPE);
            IDA_SALARY_TOTAL_DETAIL.SetSelectParamValue("W_PERSON_ID", pPERSON_ID);
            IDA_SALARY_TOTAL_DETAIL.Fill();
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
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0530_Load(object sender, EventArgs e)
        {
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            DefaultCorporation();              //Default Corp.

            W2_RB_ALLOWANCE_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W2_ALLOWANCE_TYPE.EditValue = W2_RB_ALLOWANCE_A.RadioCheckedString;
        }

        private void PERSON_JOIN_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("JOIN", string.Empty, PERSON_JOIN_COUNT.PromptText);
            }
        }

        private void PERSON_RETIRE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("RETIRE", string.Empty, PERSON_RETIRE_COUNT.PromptText);
            }
        }

        private void ADMINISTRATIVE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("ADMINISTRATIVE", string.Empty, ADMINISTRATIVE_COUNT.PromptText);
            }
        }

        private void PERSON_PROMOTION_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PROMOTION", string.Empty, PERSON_PROMOTION_COUNT.PromptText);
            }
        }

        private void NO_PAY_MASTER_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PAY_TYPE", string.Empty, NO_PAY_MASTER_COUNT.PromptText);
            }
        }

        private void NG_PAY_TYPE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PAY_TYPE", string.Empty, NG_PAY_TYPE_COUNT.PromptText);
            }
        }

        private void NO_PAY_ITEM_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PAY_ITEM", string.Empty, NO_PAY_ITEM_COUNT.PromptText);
            }
        }

        private void NO_BANK_ACCOUNTS_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("BANK_ACCOUNTS", string.Empty, NO_BANK_ACCOUNTS_COUNT.PromptText);
            }
        }

        private void NO_PAY_PROVIDE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PAY_PROVIDE", string.Empty, NO_PAY_PROVIDE_COUNT.PromptText);
            }
        }

        private void NO_BONUS_PROVIDE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("BONUS_PROVIDE", string.Empty, NO_BONUS_PROVIDE_COUNT.PromptText);
            }
        }

        private void NO_YEAR_PROVIDE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("YEAR_PROVIDE", string.Empty, NO_YEAR_PROVIDE_COUNT.PromptText);
            }
        }

        private void NO_HIRE_INSUR_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("HIRE_INSUR", string.Empty, NO_HIRE_INSUR_COUNT.PromptText);
            }
        }

        private void NO_PENSION_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("PENSION_INSUR", string.Empty, NO_PENSION_COUNT.PromptText);
            }
        }

        private void NO_MEDIC_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("MEDIC_INSUR", string.Empty, NO_MEDIC_COUNT.PromptText);
            }
        }

        private void GAP_MONTH_CLOSED_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("MONTH_DUTY", string.Empty, GAP_MONTH_CLOSED_COUNT.PromptText);
            }
        }

        private void PERSON_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail(string.Empty, string.Empty, string.Empty);
            }
        }

        private void NO_MONTH_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("NO_MONTH_DUTY", string.Empty, NO_MONTH_COUNT.PromptText);
            }
        }

        private void MONTH_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail(string.Empty, string.Empty, string.Empty);
            }
        }

        private void MONTH_CLOSED_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail(string.Empty, string.Empty, string.Empty);
            }
        }

        private void NO_SALARY_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("NO_SALARY", string.Empty, NO_SALARY_COUNT.PromptText);
            }
        }

        private void GAP_REAL_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("GAP_REAL", string.Empty, GAP_REAL_COUNT.PromptText);
            }
        }

        private void GAP_ALLOWANCE_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("GAP_ALLOWANCE", string.Empty, GAP_ALLOWANCE_COUNT.PromptText);
            }
        }

        private void GAP_DEDUCTION_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("GAP_DEDUCTION", string.Empty, GAP_DEDUCTION_COUNT.PromptText);
            }
        }

        private void REAL_MINUS_COUNT_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_DB_Detail("REAL_MINUS", string.Empty, REAL_MINUS_COUNT.PromptText);
            }
        }

        private void IGR_SALARY_ITEM_CellDoubleClick(object pSender)
        {
            if (IGR_SALARY_ITEM.RowIndex >= 0)
            {
                Search_DB_Detail(iConv.ISNull(IGR_SALARY_ITEM.GetCellValue("ITEM_TYPE")), 
                                iConv.ISNull(IGR_SALARY_ITEM.GetCellValue("ITEM_CODE")), 
                                iConv.ISNull(IGR_SALARY_ITEM.GetCellValue("ITEM_DESC")));
            }
        }

        private void IGR_SALARY_ITEM_AMOUNT_CellDoubleClick(object pSender)
        {
            if (IGR_SALARY_ITEM_AMOUNT.RowIndex >= 0)
            {
                Search_DB_Item_Detail(iConv.ISNull(IGR_SALARY_ITEM_AMOUNT.GetCellValue("ITEM_TYPE")),
                                    iConv.ISNumtoZero(IGR_SALARY_ITEM_AMOUNT.GetCellValue("ITEM_ID")),
                                    iConv.ISNull(IGR_SALARY_ITEM_AMOUNT.GetCellValue("ITEM_DESC")));
            }            
        }

        private void W2_RB_ALLOWANCE_A_CheckChanged(object sender, EventArgs e)
        {
            if (W2_RB_ALLOWANCE_A.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W2_ALLOWANCE_TYPE.EditValue = W2_RB_ALLOWANCE_A.RadioCheckedString;
            }
        }

        private void W2_RB_ALLOWANCE_D_CheckChanged(object sender, EventArgs e)
        {
            if (W2_RB_ALLOWANCE_D.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W2_ALLOWANCE_TYPE.EditValue = W2_RB_ALLOWANCE_D.RadioCheckedString;
            }
        }

        private void REAL_MINUS_COUNT_DoubleClick(object sender, EventArgs e)
        {
            Search_DB_Detail("REAL_MINUS", string.Empty, REAL_MINUS_COUNT.PromptText);
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_SALARY_TOTAL_AMOUNT_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                Search_DB_Toal_Detail(W2_ALLOWANCE_TYPE.EditValue, -1, "");
                return;
            }
            Search_DB_Toal_Detail(W2_ALLOWANCE_TYPE.EditValue, pBindingManager.DataRow["PERSON_ID"], string.Format("{0} Check", pBindingManager.DataRow["NAME"]));
        }

        #endregion

    }
}