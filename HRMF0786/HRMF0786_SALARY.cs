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

namespace HRMF0786
{
    public partial class HRMF0786_SALARY : Office2007Form
    {       

        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        #endregion;

        #region ----- Constructor -----

        public HRMF0786_SALARY(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }


        public HRMF0786_SALARY(ISAppInterface pAppInterface, object pOFFICE_TAX_NO, object pOFFICE_TAX_ID
                                , object pSTD_YYYYMM, object pPAY_YYYYMM, object pPAY_SUPPLY_DATE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            OFFICE_TAX_ID.EditValue = pOFFICE_TAX_ID;
            OFFICE_TAX_NO.EditValue = pOFFICE_TAX_NO;

            STD_YYYYMM.EditValue = pSTD_YYYYMM;
            PAY_YYYYMM.EditValue = pPAY_YYYYMM;
            PAY_SUPPLY_DATE.EditValue = pPAY_SUPPLY_DATE;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            IGR_SALARY_ITEM.LastConfirmChanges();
            IDA_SALARY_ITEM.OraSelectData.AcceptChanges();
            IDA_SALARY_ITEM.Refillable = true; 
            
            IDA_SALARY_ITEM.Fill();          
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

        private void HRMF0786_SALARY_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0786_SALARY_Shown(object sender, EventArgs e)
        {
            IDA_SALARY_ITEM.FillSchema(); 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void BTN_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_DB();
        }

        private void BTN_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_SALARY_ITEM.AddUnder();

            IGR_SALARY_ITEM.SetCellValue("OFFICE_TAX_ID", OFFICE_TAX_ID.EditValue);
            IGR_SALARY_ITEM.SetCellValue("STD_YYYYMM", STD_YYYYMM.EditValue);
            IGR_SALARY_ITEM.SetCellValue("PAY_YYYYMM", PAY_YYYYMM.EditValue);

            IGR_SALARY_ITEM.Focus(); 
        }
        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            OFFICE_TAX_NO.Focus();

            IDA_SALARY_ITEM.Update(); 
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {            
            IDA_SALARY_ITEM.Cancel(); 
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_OFFICE_TAX_DOC_ITEM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OFFICE_TAX_DOC_ITEM.SetLookupParamValue("P_TAX_FREE_YN", "N");
        }

        private void ILA_OFFICE_TAX_DOC_ITEM_TAX_FREE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OFFICE_TAX_DOC_ITEM.SetLookupParamValue("P_TAX_FREE_YN", "Y");
        }

        #endregion

        private void IDA_SALARY_ITEM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["OFFICE_TAX_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(OFFICE_TAX_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                OFFICE_TAX_NO.Focus();
                return;
            }

            if (iConv.ISNull(e.Row["STD_YYYYMM"]) == String.Empty)
            {
                int vIDX_Col = IGR_SALARY_ITEM.GetColumnToIndex("STD_YYYYMM");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_SALARY_ITEM, vIDX_Col))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                IGR_SALARY_ITEM.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["PAY_YYYYMM"]) == String.Empty)
            {
                int vIDX_Col = IGR_SALARY_ITEM.GetColumnToIndex("PAY_YYYYMM");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_SALARY_ITEM, vIDX_Col))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                IGR_SALARY_ITEM.Focus();
                return;
            }

            if (iConv.ISNull(e.Row["OFFICE_TAX_DOC_ITEM"]) == String.Empty)
            {
                int vIDX_Col = IGR_SALARY_ITEM.GetColumnToIndex("OFFICE_TAX_DOC_ITEM_NAME");
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Grid_Prompt(IGR_SALARY_ITEM, vIDX_Col))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                IGR_SALARY_ITEM.Focus();
                return;
            } 
        }
    }
}