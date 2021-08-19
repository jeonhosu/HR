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

namespace HRMF0571
{
    public partial class HRMF0571 : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----
        public HRMF0571()
        {
            InitializeComponent();
        }

        public HRMF0571(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;


        #region ----- Private Methods ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront(); 
        }

        private void Search_DB()
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YYYYMM.Focus();
                return;
            }

            IGR_ALLOWANCE_OP.LastConfirmChanges();
            IDA_ALLOWANCE_OP.OraSelectData.AcceptChanges();
            IDA_ALLOWANCE_OP.Refillable = true;

            IDA_ALLOWANCE_OP.Fill();
            IGR_ALLOWANCE_OP.Focus();
        }
         
        #endregion;

        #region ---- Default Value Setting ----

        private void Insert_DB()
        {
            IGR_ALLOWANCE_OP.SetCellValue("PERIOD_NAME", W_YYYYMM.EditValue);

            //DEFAULT OP RATE 설정 
            IDC_GET_ALLOWANCE_OP_RATE_P.ExecuteNonQuery();
            IGR_ALLOWANCE_OP.SetCellValue("OP_RATE_CODE", IDC_GET_ALLOWANCE_OP_RATE_P.GetCommandParamValue("O_OP_RATE_CODE"));
            IGR_ALLOWANCE_OP.SetCellValue("ALLOWANCE_OP_RATE_NAME", IDC_GET_ALLOWANCE_OP_RATE_P.GetCommandParamValue("O_OP_RATE_NAME"));
            IGR_ALLOWANCE_OP.SetCellValue("ALLOWANCE_OP_RATE", IDC_GET_ALLOWANCE_OP_RATE_P.GetCommandParamValue("O_OP_RATE"));
            IGR_ALLOWANCE_OP.CurrentCellMoveTo(IGR_ALLOWANCE_OP.GetColumnToIndex("PERSON_NUM"));
            IGR_ALLOWANCE_OP.Focus();
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
                    IDA_ALLOWANCE_OP.AddOver();
                    Insert_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_ALLOWANCE_OP.AddUnder();
                    Insert_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ALLOWANCE_OP.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_ALLOWANCE_OP.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    IDA_ALLOWANCE_OP.Delete();
                }
            }
        }

        #endregion;

        #region ----- FORM EVENT -----

        private void HRMF0571_Load(object sender, EventArgs e)
        {
            DefaultCorporation();
            // Default_Floor(); 
        }

        private void HRMF0571_Shown(object sender, EventArgs e)
        {
            W_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);

            RB_APPR_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_APPROVE_STATUS.EditValue = RB_APPR_NO.RadioCheckedString;

            BTN_CANCEL.Enabled = false;
            BTN_OK.Enabled = true;
             
            System.Windows.Forms.Cursor.Current = Cursors.Default;

            IDA_ALLOWANCE_OP.FillSchema();
        }

        private void irbALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            W_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();  // 버튼 상태 변경.
            IDA_ALLOWANCE_OP.Fill();
        }

        private void Default_Floor()
        {
            //작업장
            idcDEFAULT_FLOOR.ExecuteNonQuery();
            W_FLOOR_NAME.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            W_FLOOR_ID.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iString.ISNull(W_APPROVE_STATUS.EditValue);
            if (/*mAPPROVE_STATE == String.Empty ||*/ mAPPROVE_STATE == "N")
            {
                BTN_CANCEL.Enabled = false;
                BTN_OK.Enabled = true;

            }
            else if (/*mAPPROVE_STATE == String.Empty || */mAPPROVE_STATE == "A")
            {
                BTN_CANCEL.Enabled = true;
                BTN_OK.Enabled = false;

            }
            else
            {
                BTN_CANCEL.Enabled = false;
                BTN_OK.Enabled = false;
            }
        }

        //승인요청
        private void BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ALLOWANCE_OP.Update();

            IDC_SET_ALLOWANCE_OP_REQUEST.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_ALLOWANCE_OP_REQUEST.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_ALLOWANCE_OP_REQUEST.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            IDA_ALLOWANCE_OP.Fill();
        }

        //승인요청취소
        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ALLOWANCE_OP.Cancel();

            IDC_SET_ALLOWANCE_OP_REQ_CANCEL.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SET_ALLOWANCE_OP_REQ_CANCEL.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SET_ALLOWANCE_OP_REQ_CANCEL.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            IDA_ALLOWANCE_OP.Fill();
        }

        #endregion

        #region ---- LOOKUP EVENT ----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaYYYYMM_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_PERIOD_NAME", W_YYYYMM.EditValue);
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_PERIOD_NAME", IGR_ALLOWANCE_OP.GetCellValue("PERIOD_NAME"));
        }

        private void ILA_APPROVE_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_ALLOWANCE_OP_RATE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ALLOWANCE_OP_RATE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- ADAPTER EVENT -----

        private void IDA_ALLOWANCE_OP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERIOD_NAME"]) == string.Empty)
            { 
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_ALLOWANCE_OP, IGR_ALLOWANCE_OP.GetColumnToIndex("PERIOD_NAME"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_ALLOWANCE_OP, IGR_ALLOWANCE_OP.GetColumnToIndex("PERSON_NAME"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
        }

        #endregion

    }
}