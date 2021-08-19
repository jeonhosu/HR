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

namespace HRMF0707
{
    public partial class HRMF0707_FILE : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        // 입력 암호 리턴.
        public object Get_Encrypt_Password
        {
            get
            {
                return ENCRYPT_PWD.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public HRMF0707_FILE(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
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

        private void HRMF0707_FILE_Load(object sender, EventArgs e)
        {
        }

        private void HRMF0707_FILE_Shown(object sender, EventArgs e)
        {
            ENCRYPT_PWD.EditValue = string.Empty;
            CHK_ENCRYPT_PWD.EditValue = string.Empty;

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            ENCRYPT_PWD.EditValue = String.Empty;
            CHK_ENCRYPT_PWD.EditValue = String.Empty;

            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //입력함호 동일 여부 체크.
            if (iString.ISNull(ENCRYPT_PWD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ENCRYPT_PWD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ENCRYPT_PWD.Focus();
                return;
            }
            if (iString.ISNull(ENCRYPT_PWD.EditValue).Length < 8)
            {
                MessageBoxAdv.Show("암호화 비밀번호 길이는 8자리 이상이어야 합니다", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ENCRYPT_PWD.Focus();
                return;
            }
            if (iString.ISNull(ENCRYPT_PWD.EditValue).Length > 100)
            {
                MessageBoxAdv.Show("암호화 비밀번호 길이는 100자리보다 길수 없습니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ENCRYPT_PWD.Focus();
                return;
            }
            if (iString.ISNull(CHK_ENCRYPT_PWD.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CHK_ENCRYPT_PWD))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CHK_ENCRYPT_PWD.Focus();
                return;
            }
            if (iString.ISNull(ENCRYPT_PWD.EditValue) != iString.ISNull(CHK_ENCRYPT_PWD.EditValue))
            {
                MessageBoxAdv.Show("암호화 비밀번호가 다릅니다. 확인하세요.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ENCRYPT_PWD.Focus();
                return;
            }
            
            DialogResult = DialogResult.OK;
            this.Close();
        }
        
        #endregion

        
        #region ------ Lookup Event ------

        #endregion

        #region ------ Adapter Event ------

        private void idaINTEREST_RATE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["INTEREST_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10291"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaINTEREST_RATE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion             

    }
}