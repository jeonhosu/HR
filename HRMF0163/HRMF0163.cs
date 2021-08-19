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

namespace HRMF0163
{
    public partial class HRMF0163 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0163()
        {
            InitializeComponent();
        }

        public HRMF0163(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            IGR_DUTY.LastConfirmChanges();
            IDA_DUTY.OraSelectData.AcceptChanges();
            IDA_DUTY.Refillable = true;

            IDA_DUTY.Fill();
            IGR_DUTY.Focus();
        }

        private void SET_INSERT_ALLOWANCE()
        {
            IGR_DUTY.SetCellValue("ENABLED_FLAG", "Y");
            IGR_DUTY.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
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
                    if (IDA_DUTY.IsFocused)
                    {
                        IDA_DUTY.AddOver();
                        SET_INSERT_ALLOWANCE();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DUTY.IsFocused)
                    {
                        IDA_DUTY.AddUnder();
                        SET_INSERT_ALLOWANCE();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_DUTY.Update(); 
                    }
                    catch
                    {

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DUTY.IsFocused)
                    { 
                        IDA_DUTY.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DUTY.IsFocused)
                    {
                        if (IDA_DUTY.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_DUTY.Delete();
                        }
                    } 
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0163_Load(object sender, EventArgs e)
        {
            IDA_DUTY.FillSchema(); 
        }

        private void W_DUTY_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void W_DUTY_NAME_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        #endregion

        #region ---- Lookup Event -----
          
        private void ILA_DUTY_ALLOWANCE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_ALLOWANCE_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_DUTY_DAY_LEAVE_CONTROL_FLAG_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_DAY_LEAVE_CONTROL_FLAG");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_DUTY_ATTEND_FLAG_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_ATTEND_FLAG");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        } 

        private void ILA_DUTY_GROUP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["DUTY_CODE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DUTY, IGR_DUTY.GetColumnToIndex("DUTY_CODE")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["DUTY_NAME"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DUTY, IGR_DUTY.GetColumnToIndex("DUTY_NAME")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_DUTY, IGR_DUTY.GetColumnToIndex("EFFECTIVE_DATE_FR")))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
        }

        #endregion

    }
}