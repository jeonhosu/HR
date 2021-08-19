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

namespace HRMF0431
{
    public partial class HRMF0431 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #region ----- Constructor -----
        public HRMF0431(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void SEARCH_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (W_YEAR_YYYY.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IGR_YEAR_INSUR_MEDIC.LastConfirmChanges();
            IDA_YEAR_INSUR_MEDIC.OraSelectData.AcceptChanges();
            IDA_YEAR_INSUR_MEDIC.Refillable = true;

            IDA_YEAR_INSUR_MEDIC.Fill();
            IGR_YEAR_INSUR_MEDIC.Focus();
        }
         
        private void InsertDB()
        {
            IGR_YEAR_INSUR_MEDIC.Focus();
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

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_YEAR_INSUR_MEDIC.IsFocused)
                    {
                        IDA_YEAR_INSUR_MEDIC.AddOver();
                        InsertDB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_YEAR_INSUR_MEDIC.IsFocused)
                    {
                        IDA_YEAR_INSUR_MEDIC.AddUnder();
                        InsertDB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_YEAR_INSUR_MEDIC.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_YEAR_INSUR_MEDIC.Cancel(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if(IDA_YEAR_INSUR_MEDIC.CurrentRow.RowState == DataRowState.Added)
                    {
                        IDA_YEAR_INSUR_MEDIC.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0431_Load(object sender, EventArgs e)
        { 
        }

        private void HRMF0431_Shown(object sender, EventArgs e)
        {
            W_YEAR_YYYY.EditValue = iDate.ISYear(DateTime.Today);
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]

            DefaultCorporation();                  // Corp Default Value Setting. 
            // FillSchema
            IDA_YEAR_INSUR_MEDIC.FillSchema();
        }

        private void BTN_OPEN_DATE_Y_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iConv.ISNull(V_OPEN_DATE.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_OPEN_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_OPEN_DATE.Focus();
                return;
            }

            IDC_EXEC_OPEN_DATE.SetCommandParamValue("P_OPEN_STATUS", "OPEN_YES");
            IDC_EXEC_OPEN_DATE.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_EXEC_OPEN_DATE.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_EXEC_OPEN_DATE.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            V_OPEN_DATE.EditValue = DBNull.Value;
            SEARCH_DB();
        }

        private void BTN_OPEN_DATE_C_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDC_EXEC_OPEN_DATE.SetCommandParamValue("P_OPEN_STATUS", "OPEN_CANCEL"); 
            IDC_EXEC_OPEN_DATE.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_EXEC_OPEN_DATE.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_EXEC_OPEN_DATE.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            V_OPEN_DATE.EditValue = DBNull.Value;
            SEARCH_DB();
        }

        private void BTN_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0431_EXPORT vHRMF0431_EXPORT = new HRMF0431_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                                , W_YEAR_YYYY.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0431_EXPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0431_EXPORT.ShowDialog();
            vHRMF0431_EXPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                SEARCH_DB();
            }
            vHRMF0431_EXPORT.Dispose();
        }


        private void BTN_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0431_IMPORT vHRMF0431_IMPORT = new HRMF0431_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, W_CORP_ID.EditValue, W_YEAR_YYYY.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0431_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0431_IMPORT.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            vHRMF0431_IMPORT.Dispose();
            SEARCH_DB();
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void IDA_YEAR_INSUR_MEDIC_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["YEAR_YYYY"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("HRM_10009"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void IDA_YEAR_INSUR_MEDIC_PreDelete(ISPreDeleteEventArgs e)
        {
             
        } 

        #endregion

        #region ----- Lookup Event -----

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y"); 
        }

        private void ILA_POST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        } 

        private void ilaPERSON_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", W_DEPT_ID.EditValue);
            ildPERSON.SetLookupParamValue("W_POST_ID", W_POST_ID.EditValue);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", W_FLOOR_ID.EditValue);
        }

        private void ILA_PERSON_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_POST_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", DBNull.Value);
        }
        #endregion

    }
}