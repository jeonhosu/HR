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

namespace HRMF0799
{
    public partial class HRMF0799 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0799(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

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
            ILD_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void SEARCH_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(W_CLOSING_YYYY.EditValue) == string.Empty)
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10031"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            IDA_YEAR_CLOSING.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_YEAR_CLOSING.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_YEAR_CLOSING.Fill();

            IGR_YEAR_CLOSING.Focus();
        }

        #endregion

        #region ----- Application_MainButtonClick -----
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    //if (IDA_YEAR_CLOSING.IsFocused)
                    //{
                    //    IDA_YEAR_CLOSING.AddOver();
                    //    IGR_YEAR_CLOSING.SetCellValue("CORP_ID", W_CORP_ID.EditValue);
                    //    IGR_YEAR_CLOSING.SetCellValue("CLOSING_YYYY", W_CLOSING_YYYY.EditValue);

                    //    IGR_YEAR_CLOSING.CurrentCellMoveTo(IGR_YEAR_CLOSING.GetColumnToIndex("CLOSING_YYYY"));
                    //    IGR_YEAR_CLOSING.Focus();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    //if (IDA_YEAR_CLOSING.IsFocused)
                    //{
                    //    IDA_YEAR_CLOSING.AddUnder();
                    //    IGR_YEAR_CLOSING.SetCellValue("CORP_ID", W_CORP_ID.EditValue);
                    //    IGR_YEAR_CLOSING.SetCellValue("CLOSING_YYYY", W_CLOSING_YYYY.EditValue);

                    //    IGR_YEAR_CLOSING.CurrentCellMoveTo(IGR_YEAR_CLOSING.GetColumnToIndex("CLOSING_YYYY"));
                    //    IGR_YEAR_CLOSING.Focus();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_YEAR_CLOSING.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_YEAR_CLOSING.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                   
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0799_Load(object sender, EventArgs e)
        { 
            IDA_YEAR_CLOSING.FillSchema();
            DefaultCorporation(); 

            W_CLOSING_YYYY.EditValue = iDate.ISYear(DateTime.Today);
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ILD_COMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ILD_COMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }

        private void ibtCREATION_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string sStatus = "";
            string sMessage = null;
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_CLOSING_YYYY.EditValue) == string.Empty)
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10031"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CLOSING_YYYY.Focus();
                return;
            }
            IDC_CLOSING_CREATE.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDC_CLOSING_CREATE.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDC_CLOSING_CREATE.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID); 
            IDC_CLOSING_CREATE.ExecuteNonQuery();
            sStatus = iString.ISNull(IDC_CLOSING_CREATE.GetCommandParamValue("O_STATUS"));
            sMessage = iString.ISNull(IDC_CLOSING_CREATE.GetCommandParamValue("O_MESSAGE"));
            if(sStatus == "F")
            {
                if(sMessage != string.Empty)
                {
                    MessageBoxAdv.Show(sMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            MessageBoxAdv.Show(sMessage, "Closing Create");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_YEAR_CLOSING_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        private void IDA_YEAR_CLOSING_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_YEAR_CLOSING_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_CLOSING");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}