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

namespace HRMF0539
{
    public partial class HRMF0539 : Office2007Form
    {

        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        Object mSESSION_ID;

        #endregion;

        #region ----- Constructor -----

        public HRMF0539(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

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

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
             
            IDA_PAYSLIP_EXCEPTION.Fill();
            IGR_YEAR_MANAGEMENT.Focus(); 
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

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        

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
                    IDA_PAYSLIP_EXCEPTION.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_PAYSLIP_EXCEPTION.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if(IDA_PAYSLIP_EXCEPTION.IsFocused)
                    {
                        IDA_PAYSLIP_EXCEPTION.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0539_Load(object sender, EventArgs e)
        {
       
        }

        private void HRMF0539_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();       //Default Corp.
            W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);

            V_STATUS_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_STATUS.EditValue = V_STATUS_ALL.RadioCheckedString;

            BTN_EXCEL_EXPORT.BringToFront();
            BTN_EXCEL_IMPORT.BringToFront();


            //IDC_GET_SESSION_ID_P.ExecuteNonQuery();
            //mSESSION_ID = IDC_GET_SESSION_ID_P.GetCommandParamValue("O_SESSION_ID");
        }

        private void V_STATUS_ALL_Click(object sender, EventArgs e)
        {
            W_STATUS.EditValue = V_STATUS_ALL.RadioCheckedString;
        }

        private void V_STATUS_YES_Click(object sender, EventArgs e)
        {
            W_STATUS.EditValue = V_STATUS_YES.RadioCheckedString;
        }

        private void V_STATUS_NO_Click(object sender, EventArgs e)
        {
            W_STATUS.EditValue = V_STATUS_NO.RadioCheckedString;
        }
         
        #endregion  

        #region ----- Adapter Event -----
        // Allowance 항목.
        private void idaADD_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }  
        }

        private void idaADD_ALLOWANCE_PreDelete(ISPreDeleteEventArgs e)
        {
        }   
        #endregion

        #region ----- LookUp Event -----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today,2)));
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_POST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DEPT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y"); 
        }

        #endregion

        private void BTN_EXCEL_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0539_EXPORT vHRMF0539_EXPORT = new HRMF0539_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                                , W_STD_YYYYMM.EditValue 
                                                                , W_POST_ID.EditValue, W_POST_NAME.EditValue
                                                                , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                                , W_FLOOR_ID.EditValue, W_FLOOR_NAME.EditValue
                                                                , W_PERSON_ID.EditValue, W_PERSON_NUM.EditValue, W_PERSON_NAME.EditValue); 
            vdlgResult = vHRMF0539_EXPORT.ShowDialog();
            vHRMF0539_EXPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void BTN_EXCEL_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0539_IMPORT vHRMF0539_IMPORT = new HRMF0539_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, W_CORP_ID.EditValue
                                                                , W_STD_YYYYMM.EditValue, mSESSION_ID); 
            vdlgResult = vHRMF0539_IMPORT.ShowDialog();
            vHRMF0539_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }
    }
}