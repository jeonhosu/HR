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

namespace HRMF0335
{
    public partial class HRMF0335 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0335(Form pMainForm, ISAppInterface pAppInterface)
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
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE.Focus();
                return;
            }

            //object vPERSON_ID = IGR_WORK_EXTEND.GetCellValue("PERSON_ID");
            IDA_WORK_EXTEND.Fill();
            IGR_WORK_EXTEND.Focus();

            //if (iString.ISNull(vPERSON_ID) != string.Empty)
            //{
            //    int vIDX_PERSON_ID = IGR_WORK_EXTEND.GetColumnToIndex("PERSON_ID");
            //    for (int i = 0; i < IGR_WORK_EXTEND.RowCount; i++)
            //    {
            //        if (iString.ISNull(vPERSON_ID) == iString.ISNull(IGR_WORK_EXTEND.GetCellValue(i, vIDX_PERSON_ID)))
            //        {
            //            IGR_WORK_EXTEND.CurrentCellMoveTo(i, 1);
            //            return;
            //        }
            //    }
            //    IGR_WORK_EXTEND.CurrentCellMoveTo(1, 1);
            //}
        }

        private void Insert_Default_Value()
        {
           // IGR_WORK_EXTEND.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            IGR_WORK_EXTEND.SetCellValue("WORK_DATE", WORK_DATE.EditValue);
          //  IGR_WORK_EXTEND.SetCellValue("ENABLED_YN", "Y");
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
                    //if (IDA_WORK_EXTEND.IsFocused)
                    //{
                    //    IDA_WORK_EXTEND.AddOver();
                    //    Insert_Default_Value();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (IDA_WORK_EXTEND.IsFocused)
                    //{
                    //    IDA_WORK_EXTEND.AddUnder();
                    //    Insert_Default_Value();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_WORK_EXTEND.IsFocused)
                    {
                        IDA_WORK_EXTEND.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_WORK_EXTEND.IsFocused)
                    {
                        IDA_WORK_EXTEND.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0335_Load(object sender, EventArgs e)
        {
            IDA_WORK_EXTEND.FillSchema();
        }

        private void HRMF0335_Shown(object sender, EventArgs e)
        {
            WORK_DATE.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
        }


        private void BTN_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0335_UPLOAD vHRMF0335_UPLOAD = new HRMF0335_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0335_UPLOAD.ShowDialog();
            vHRMF0335_UPLOAD.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }            
        }
        #endregion  

        #region ----- Adapter Event ------

        private void IDA_WORK_SCHEDULE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WORK_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["DUTY_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(DUTY_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["ENABLED_YN"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", "Enabled Y/N")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

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

        private void igbCONDITION_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}