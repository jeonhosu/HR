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

namespace HRMF0393
{
    public partial class HRMF0393 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0393(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_START_DATE.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_START_DATE.Focus();
                return;
            }


            IDA_DEVICE_LOG_M.Fill();
            IGR_DEVICE_LOG.Focus(); 
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
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DEVICE_LOG_M.IsFocused)
                    {
                        IDA_DEVICE_LOG_M.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (IDA_DEVICE_LOG.IsFocused)
                    //{
                    //    IDA_DEVICE_LOG.Delete();
                    //}
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0393_Load(object sender, EventArgs e)
        {
            W_CORP_NAME.BringToFront();
            irbALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_STATUS_FLAG.EditValue = irbALL.RadioCheckedString;

            igbSET_INTERFACE.Visible = false;
        }

        private void HRMF0393_Shown(object sender, EventArgs e)
        {
            W_START_DATE.EditValue = DateTime.Today;
            W_END_DATE.EditValue = W_START_DATE.EditValue;

            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]

            IDA_DEVICE_LOG_M.FillSchema();
        }

        private void WORK_DATE_FR_0_EditValueChanged(object pSender)
        {
            W_END_DATE.EditValue = W_START_DATE.EditValue;
        }

        private void irbALL_Click(object sender, EventArgs e)
        {
            if(irbALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS_FLAG.EditValue = irbALL.RadioCheckedString;
            }
        }

        private void irbDUTY_Click(object sender, EventArgs e)
        {
            if (irbDUTY.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS_FLAG.EditValue = irbDUTY.RadioCheckedString;
            }
        }

        private void irbFOOD_Click(object sender, EventArgs e)
        {
            if (irbFOOD.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_STATUS_FLAG.EditValue = irbFOOD.RadioCheckedString;
            }
        }

        private void BTN_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0393_UPLOAD vHRMF0393_UPLOAD = new HRMF0393_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0393_UPLOAD.ShowDialog();
            vHRMF0393_UPLOAD.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }            
        }

        #endregion  

        #region ----- Adapter Event ------
         

        #endregion

        #region ----- LookUp Event -----
        
        private void ILA_DUTY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_HOLY_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_CLASS");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}