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

namespace HRMF0225
{
    public partial class HRMF0225 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0225()
        {
            InitializeComponent();
        }

        public HRMF0225(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultValues()
        {
            // Lookup SETTING
            ILD_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            if (iConv.ISNull(STAT_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            if (iConv.ISNull(END_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            if (Convert.ToDateTime(STAT_DATE.EditValue) > Convert.ToDateTime(END_DATE.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STAT_DATE.Focus();
                return;
            }

            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                IDA_LICENSE.Fill();
                IGR_LICENSE.Focus();
            }
            else if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                IDA_FOREIGN_LANGUAGE.Fill();
                IGA_FOREIGN_LANGUAGE.Focus();
            }
        }
        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
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
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ------ Lookup event ------

        private void ILA_POST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("POST", "Y");
        }

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }


        private void ILA_PERSON_SelectedRowData(object pSender)
        {
            Search_DB();
        }

        #endregion;

        #region ----Form event -----

        private void HRMF0225_Load(object sender, EventArgs e)
        {
            IDA_LICENSE.FillSchema();
        }

        private void HRMF0225_Shown(object sender, EventArgs e)
        {
            DefaultValues();

            STAT_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = DateTime.Today;
        }

        #endregion

        private void PERSON_NUM_Load(object sender, EventArgs e)
        {

        }
    }
}