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

namespace HRMF0162
{
    public partial class HRMF0162 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0162()
        {
            InitializeComponent();
        }

        public HRMF0162(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            IGR_OT_TYPE.LastConfirmChanges();
            IDA_OT_TYPE.OraSelectData.AcceptChanges();
            IDA_OT_TYPE.Refillable = true;

            IDA_OT_TYPE.Fill();
            IGR_OT_TYPE.Focus();
        }

        private void SET_INSERT_ALLOWANCE()
        {
            IGR_OT_TYPE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_OT_TYPE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
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
                    if (IDA_OT_TYPE.IsFocused)
                    {
                        IDA_OT_TYPE.AddOver();
                        SET_INSERT_ALLOWANCE();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_OT_TYPE.IsFocused)
                    {
                        IDA_OT_TYPE.AddUnder();
                        SET_INSERT_ALLOWANCE();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_OT_TYPE.Update(); 
                    }
                    catch
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OT_TYPE.IsFocused)
                    { 
                        IDA_OT_TYPE.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OT_TYPE.IsFocused)
                    {
                        if (IDA_OT_TYPE.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_OT_TYPE.Delete();
                        }
                    } 
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0162_Load(object sender, EventArgs e)
        {
            IDA_OT_TYPE.FillSchema(); 
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
         
        private void ILA_ALLOWANCE_ITEM_10_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
         
        private void ILA_ALLOWANCE_ITEM_20_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CHG_OT_TYPE_H_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OT_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CHG_OT_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OT_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CHG_OT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OT_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CHG_OT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OT_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_OT_GROUP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OT_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["OT_TYPE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_OT_TYPE, IGR_OT_TYPE.GetColumnToIndex("OT_TYPE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["OT_TYPE_NAME"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_OT_TYPE, IGR_OT_TYPE.GetColumnToIndex("OT_TYPE_NAME")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_OT_TYPE, IGR_OT_TYPE.GetColumnToIndex("EFFECTIVE_DATE_FR")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
        }

        #endregion

    }
}