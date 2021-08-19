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

namespace HRMF0211
{
    public partial class HRMF0211 : Office2007Form
    {
        #region ----- Constructor -----
        public HRMF0211(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Property / Method -----

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
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(STD_DATE_0.EditValue.ToString()))
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_EMPLOYE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_EMPLOYE.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_EMPLOYE.Fill();
            igrEMPLOYE.Focus();
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----
        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_EMPLOYE.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_EMPLOYE.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_EMPLOYE.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_EMPLOYE.IsFocused)
                    {
                        IDA_EMPLOYE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_EMPLOYE.IsFocused)
                    {
                        IDA_EMPLOYE.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0211_Load(object sender, EventArgs e)
        {
            this.Visible = true;
            // FillSchema
            IDA_EMPLOYE.FillSchema();
            STD_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }

        #endregion

        #region ----- Data Adapter Event ----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {

        }
        #endregion

        #region ----- Lookup Event -----

        private void ilaPOST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");

            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // DEPT
            ildDEPT.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildDEPT.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // PERSON
            ildPERSON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildPERSON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

    }
}