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

namespace HRMF0224
{
    public partial class HRMF0224 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0224()
        {
            InitializeComponent();
        }

        public HRMF0224(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP_0.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP_0.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }


        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }

            if (TB_BASE.SelectedTab.TabIndex == TP_SCHOLAPSHIP.TabIndex)
            {
                IDA_SCHOLARSHIP_STATE_1.Fill();
                IGR_SCHOLARSHIP_CURRENT.Focus();
            }
            else if (TB_BASE.SelectedTab.TabIndex == TP_CAREER.TabIndex)
            {
                IDA_CAREER.Fill();
                IGR_CAREER.Focus();
            }
        }

        private void isSetCommonLookUpParameter(string P_GROUP_CODE, String P_USABLE_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_USABLE_YN);
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
            }
        }

        #endregion;

        #region ----Form event -----

        private void HRMF0224_Shown(object sender, EventArgs e)
        {
            DefaultValues();
        }
 
        #endregion


        #region ------ Lookup event ------
        private void ILA_EMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EMPLOYE_TYPE", "N");
        }

        #endregion


        #region ------ Adapter event -----


        #endregion
    }
}