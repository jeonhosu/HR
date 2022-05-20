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

namespace HRMF0161
{
    public partial class HRMF0161 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0161()
        {
            InitializeComponent();
        }

        public HRMF0161(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_DEDUCTION.TabIndex)
            {
                IGR_DEDUCTION_ACCOUNT_9.LastConfirmChanges();
                IDA_DEDUCTION_ACCOUNT_9.OraSelectData.AcceptChanges();
                IDA_DEDUCTION_ACCOUNT_9.Refillable = true;

                IGR_DEDUCTION_ACCOUNT_1.LastConfirmChanges();
                IDA_DEDUCTION_ACCOUNT_1.OraSelectData.AcceptChanges();
                IDA_DEDUCTION_ACCOUNT_1.Refillable = true;

                IGR_DEDUCTION.LastConfirmChanges();
                IDA_DEDUCTION.OraSelectData.AcceptChanges();
                IDA_DEDUCTION.Refillable = true;

                IDA_DEDUCTION.Fill();
                IGR_DEDUCTION.Focus(); 
            }
            else
            {
                IGR_ALLOWANCE_ACCOUNT_9.LastConfirmChanges();
                IDA_ALLOWANCE_ACCOUNT_9.OraSelectData.AcceptChanges();
                IDA_ALLOWANCE_ACCOUNT_9.Refillable = true;

                IGR_ALLOWANCE_ACCOUNT_1.LastConfirmChanges();
                IDA_ALLOWANCE_ACCOUNT_1.OraSelectData.AcceptChanges();
                IDA_ALLOWANCE_ACCOUNT_1.Refillable = true;

                IGR_ALLOWANCE.LastConfirmChanges();
                IDA_ALLOWANCE.OraSelectData.AcceptChanges();
                IDA_ALLOWANCE.Refillable = true;

                IDA_ALLOWANCE.Fill();
                IGR_ALLOWANCE.Focus();
            }
        }

        private void SET_INSERT_ALLOWANCE()
        {
            IGR_ALLOWANCE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_ALLOWANCE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
        }

        private void SET_INSERT_DEDUCTION()
        {
            IGR_DEDUCTION.SetCellValue("ENABLED_FLAG", "Y");
            IGR_DEDUCTION.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
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
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_ALLOWANCE.IsFocused)
                    {
                        IDA_ALLOWANCE.AddOver();
                        SET_INSERT_ALLOWANCE();
                    }
                    else if (IDA_DEDUCTION.IsFocused)
                    {
                        IDA_DEDUCTION.AddOver();
                        SET_INSERT_DEDUCTION();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ALLOWANCE.IsFocused)
                    {
                        IDA_ALLOWANCE.AddUnder();
                        SET_INSERT_ALLOWANCE();
                    }
                    else if (IDA_DEDUCTION.IsFocused)
                    {
                        IDA_DEDUCTION.AddUnder();
                        SET_INSERT_DEDUCTION();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_ALLOWANCE.Update();
                        IDA_DEDUCTION.Update();
                    }
                    catch
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ALLOWANCE.IsFocused)
                    {
                        IDA_ALLOWANCE_ACCOUNT_1.Cancel();
                        IDA_ALLOWANCE_ACCOUNT_9.Cancel();
                        IDA_ALLOWANCE.Cancel();
                    }
                    else if (IDA_ALLOWANCE_ACCOUNT_1.IsFocused)
                    {
                        IDA_ALLOWANCE_ACCOUNT_1.Cancel();
                    }
                    else if (IDA_ALLOWANCE_ACCOUNT_9.IsFocused)
                    {
                        IDA_ALLOWANCE_ACCOUNT_9.Cancel();
                    }
                    else if (IDA_ALLOWANCE_CAL_METHOD.IsFocused)
                    {
                        IDA_ALLOWANCE_CAL_METHOD.Cancel();
                    }
                    else if (IDA_DEDUCTION.IsFocused)
                    {
                        IDA_DEDUCTION_ACCOUNT_1.Cancel();
                        IDA_DEDUCTION_ACCOUNT_9.Cancel();
                        IDA_DEDUCTION.Cancel();
                    }
                    else if (IDA_DEDUCTION_ACCOUNT_1.IsFocused)
                    {
                        IDA_DEDUCTION_ACCOUNT_1.Cancel();
                    }
                    else if (IDA_DEDUCTION_ACCOUNT_9.IsFocused)
                    {
                        IDA_DEDUCTION_ACCOUNT_9.Cancel();
                    }
                    else if(IDA_DEDUCTION_CAL_METHOD.IsFocused)
                    {
                        IDA_DEDUCTION_CAL_METHOD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ALLOWANCE.IsFocused)
                    {
                        if (IDA_ALLOWANCE.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_ALLOWANCE.Delete();
                        }
                    }
                    else if(IDA_ALLOWANCE_ACCOUNT_1.IsFocused)
                    {
                        IDA_ALLOWANCE_ACCOUNT_1.Delete();
                    }
                    else if(IDA_ALLOWANCE_ACCOUNT_9.IsFocused)
                    {
                        IDA_ALLOWANCE_ACCOUNT_9.Delete();
                    }
                    else if(IDA_ALLOWANCE_CAL_METHOD.IsFocused)
                    {
                        IDA_DEDUCTION_CAL_METHOD.Delete();
                    }
                    else if (IDA_DEDUCTION.IsFocused)
                    {
                        if (IDA_DEDUCTION.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_DEDUCTION.Delete();
                        }
                    }
                    else if (IDA_DEDUCTION_ACCOUNT_1.IsFocused)
                    {
                        IDA_DEDUCTION_ACCOUNT_1.Delete();
                    }
                    else if (IDA_DEDUCTION_ACCOUNT_9.IsFocused)
                    {
                        IDA_DEDUCTION_ACCOUNT_9.Delete();
                    }
                    else if(IDA_DEDUCTION_CAL_METHOD.IsFocused)
                    {
                        IDA_DEDUCTION_CAL_METHOD.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0161_Load(object sender, EventArgs e)
        {
            IDA_ALLOWANCE.FillSchema();
            IDA_ALLOWANCE_ACCOUNT_1.FillSchema();
            IDA_ALLOWANCE_ACCOUNT_9.FillSchema();

            IDA_DEDUCTION.FillSchema();
            IDA_DEDUCTION_ACCOUNT_1.FillSchema();
            IDA_DEDUCTION_ACCOUNT_9.FillSchema();
        }

        private void W_WORK_CENTER_DESC_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        #endregion

        #region ---- Lookup Event -----

        private void ILA_ACCOUNT_CONTROL_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VENDOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
         
        private void ILA_TAX_FREE_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_TAX_FREE_TYPE.SetLookupParamValue("W_ITEM_TYPE", "A");
            ILD_TAX_FREE_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");   
        }

        private void ILA_ACCOUNT_DR_CR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        private void ILA_ACCOUNT_DR_CR_A_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_A_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_RETIRE_SALARY_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "RETIRE_SALARY_ITEM");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y"); 
        }

        private void ILA_ALLOWANCE_ITEM_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE_ITEM_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_ALLOWANCE_EXP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE_EXP_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_VENDOR_A_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_DEDUCTION_ITEM_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DEDUCTION_ITEM_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_TAX_FREE_TYPE_D_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_TAX_FREE_TYPE.SetLookupParamValue("W_ITEM_TYPE", "D");
            ILD_TAX_FREE_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_VENDOR_D_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_VENDOR_D_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_D_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_CONTROL_D_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ACCOUNT_CONTROL.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_DR_CR_D_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ACCOUNT_DR_CR_D_9_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "ACCOUNT_DR_CR");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_ALLOWANCE_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y"); 
        }

        private void ILA_DEDUCTION_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DEDUCTION");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_ALLOWANCE_CAL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE_ETC");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DEDUCTION_CAL_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DEDUCTION_ETC");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["ALLOWANCE_CODE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ALLOWANCE, IGR_ALLOWANCE.GetColumnToIndex("ALLOWANCE_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["ALLOWANCE_NAME"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ALLOWANCE, IGR_ALLOWANCE.GetColumnToIndex("ALLOWANCE_NAME")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ALLOWANCE, IGR_ALLOWANCE.GetColumnToIndex("EFFECTIVE_DATE_FR")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //if (iConv.ISNull(e.Row["ALLOWANCE_VIEW_CODE"]) == string.Empty)
            //{
            //    e.Cancel = true;
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ALLOWANCE, IGR_ALLOWANCE.GetColumnToIndex("ALLOWANCE_VIEW_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            //if (iConv.ISNull(e.Row["ALLOWANCE_PRINT_CODE"]) == string.Empty)
            //{
            //    e.Cancel = true;
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_ALLOWANCE, IGR_ALLOWANCE.GetColumnToIndex("ALLOWANCE_PRINT_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
        }

        private void IDA_ALLOWANCE_ACCOUNT_1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void IDA_ALLOWANCE_ACCOUNT_9_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void IDA_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

            if (iConv.ISNull(e.Row["DEDUCTION_CODE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEDUCTION, IGR_DEDUCTION.GetColumnToIndex("DEDUCTION_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["DEDUCTION_NAME"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEDUCTION, IGR_DEDUCTION.GetColumnToIndex("DEDUCTION_NAME")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEDUCTION, IGR_DEDUCTION.GetColumnToIndex("EFFECTIVE_DATE_FR")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //if (iConv.ISNull(e.Row["DEDUCTION_VIEW_CODE"]) == string.Empty)
            //{
            //    e.Cancel = true;
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEDUCTION, IGR_DEDUCTION.GetColumnToIndex("DEDUCTION_VIEW_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            //if (iConv.ISNull(e.Row["DEDUCTION_PRINT_CODE"]) == string.Empty)
            //{
            //    e.Cancel = true;
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DEDUCTION, IGR_DEDUCTION.GetColumnToIndex("DEDUCTION_PRINT_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
        }

        private void IDA_DEDUCTION_ACCOUNT_1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void IDA_DEDUCTION_ACCOUNT_9_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        #endregion

    }
}