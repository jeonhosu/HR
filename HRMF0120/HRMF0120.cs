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

namespace HRMF0120
{
    public partial class HRMF0120 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0120()
        {
            InitializeComponent();
        }

        public HRMF0120(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        private void Search_DB()
        {
            IDA_MASTER_NUM.Fill();
            IGR_MASTER_NUM.Focus();
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
                    if (IDA_MASTER_NUM.IsFocused)
                    {
                        IDA_MASTER_NUM.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_MASTER_NUM.IsFocused)
                    {
                        IDA_MASTER_NUM.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_MASTER_NUM.IsFocused)
                    {
                        IDA_MASTER_NUM.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_MASTER_NUM.IsFocused)
                    {
                        IDA_MASTER_NUM.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_MASTER_NUM.IsFocused)
                    {
                        if (IDA_MASTER_NUM.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_MASTER_NUM.Delete();
                        }
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0120_Load(object sender, EventArgs e)
        {
            IDA_MASTER_NUM.FillSchema();
        }

        private void IGR_MASTER_NUM_CellDoubleClick(object pSender)
        {
            if (IGR_MASTER_NUM.Row < 0)
            {
                return;
            }

            IDA_MASTER_NUM_HISTORY.Fill();

            TB_MASTER_NUM.SelectedIndex = 1;
            TB_MASTER_NUM.Focus();
        }

        #endregion

        
        #region ----- Adapter Event ----

        private void IDA_MASTER_NUM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["MASTER_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10104"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_MASTER_NUM_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ---- Lookup Event ----

        private void ILA_DATE_FORMAT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DATE_FORMAT");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}