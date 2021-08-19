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

namespace HRMF0781
{
    public partial class HRMF0781_FILE : Office2007Form
    {       

        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion

        #region ----- Constructor -----

        public HRMF0781_FILE(ISAppInterface pAppInterface, object pCORP_ID)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            CORP_ID.EditValue = pCORP_ID;
        }

        #endregion;

        #region ----- Private Methods ----

        
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
            try
            {
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
            }
            catch
            {
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
                    if (IDA_HOMETAX.IsFocused)
                    {
                        IDA_HOMETAX.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_HOMETAX.IsFocused)
                    {
                        IDA_HOMETAX.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_HOMETAX.IsFocused)
                    {
                        IDA_HOMETAX.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0781_FILE_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0781_FILE_Shown(object sender, EventArgs e)
        {
            IDA_HOMETAX.Fill();
            if (IDA_HOMETAX.OraSelectData.Rows.Count == 0)
            {
                IDA_HOMETAX.AddUnder();
            }
            HOMETAX_ID.Focus();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_HOMETAX.Update();
            if (IDA_HOMETAX.UpdateChangedRowCount != 0 || IDA_HOMETAX.CurrentRow.RowState == DataRowState.Unchanged)
            {
                DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_HOMETAX.Cancel();
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion 
        
        #region ----- Adapter Event -----

        private void IDA_HOMETAX_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel=true;
                return;
            }
            if (iString.ISNull(e.Row["HOMETAX_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(HOMETAX_ID))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ENCRYPT_PWD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ENCRYPT_PWD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ENCRYPT_PWD"]).Length < 4)
            {
                MessageBoxAdv.Show("암호화 비밀번호 길이는 4자리 이상이어야 합니다", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ENCRYPT_PWD"]).Length > 100)
            {
                MessageBoxAdv.Show("암호화 비밀번호 길이는 100자리보다 길수 없습니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CHK_ENCRYPT_PWD"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CHK_ENCRYPT_PWD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ENCRYPT_PWD"]) != iString.ISNull(e.Row["CHK_ENCRYPT_PWD"]))
            {
                MessageBoxAdv.Show("암호화 비밀번호가 다릅니다. 확인하세요.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USER_DEPT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(USER_DEPT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USER_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&USER_NAME:={0}", Get_Edit_Prompt(USER_DEPT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USER_TEL_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&USER_NAME:={0}", Get_Edit_Prompt(USER_TEL_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}