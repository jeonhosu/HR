using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace HRMF0384
{
    public partial class HRMF0384 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0384()
        {
            InitializeComponent();
        }

        public HRMF0384(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            
            WORK_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            WORK_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (WORK_CORP_ID_0.EditValue == null)
            {// 업체. - 업체정보는 필수입니다. 선택하세요.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자 - 시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                IDA_DAY_INTERFACE_SUMMARY.Fill();
                IGR_DAY_INTERFACE_SUMMARY.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                IDA_JOIN_PERSON.Fill();
                IGR_JOIN_PERSON.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 3)
            {
                IDA_RETIRE_PERSON.Fill();
                IGR_RETIRE_PERSON.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 4)
            {
                IDA_NO_WORK_PERSON.Fill();
                IGR_NO_WORK_PERSON.Focus();
            }
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
                
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
  
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
       
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DAY_INTERFACE_SUMMARY.IsFocused)
                    {
                        IDA_DAY_INTERFACE_SUMMARY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DAY_INTERFACE_SUMMARY.IsFocused)
                    {
                        IDA_DAY_INTERFACE_SUMMARY.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0384_Load(object sender, EventArgs e)
        {

        }

        private void HRMF0384_Shown(object sender, EventArgs e)
        {
            WORK_DATE_0.EditValue = DateTime.Today;
            CB_PRE_WORK_DATE_YN_0.CheckedState = ISUtil.Enum.CheckedState.Checked;

            DefaultCorporation();

            WORK_DATE_0.Focus();

            RB_WORK_CENTER.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SORT_FLAG.EditValue = "WC";

        }

        private void RB_WORK_CENTER_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;
            if (vRadio.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SORT_FLAG.EditValue = vRadio.RadioCheckedString;
            }
        }

        private void TB_MAIN_Click(object sender, EventArgs e)
        {
            //Search_DB();
        }

        #endregion

        #region ----- Lookup event -----

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        
        #endregion

    }
}