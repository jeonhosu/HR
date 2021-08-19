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



namespace HRMF0229
{
    public partial class HRMF0229 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public HRMF0229()
        {
            InitializeComponent();
        }

        public HRMF0229(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

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

            if (iConv.ISNull(START_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //e.Cancel = true;
                return;
            }


            IDA_HISTORY.Fill();
            IGR_HISTORY.Focus();
            
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
                    if (IDA_HISTORY.IsFocused)
                    {
                        IDA_HISTORY.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_HISTORY.IsFocused)
                    {
                        IDA_HISTORY.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_HISTORY.IsFocused)
                    {
                        IDA_HISTORY.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_HISTORY.IsFocused)
                    {
                        IDA_HISTORY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_HISTORY.IsFocused)
                    {
                        IDA_HISTORY.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void HRMF0229_Load(object sender, EventArgs e)
        {
            IDA_HISTORY.FillSchema();
        }

        private void HRMF0229_Shown(object sender, EventArgs e)
        {
            DefaultValues();

            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = DateTime.Today;
        }
        #endregion;

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }
    }
}