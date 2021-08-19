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

namespace HRMF0509
{
    public partial class HRMF0509 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0509()
        {
            InitializeComponent();
        }

        public HRMF0509(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
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

        private void HRMF0509_Load(object sender, EventArgs e)
        {

        }

        private void HRMF0509_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();              //Default Corp.

            PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);
        }

        #region ----- Lookup Event -----

        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }
        
        #endregion

        private void ibtEXECUTE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME.Focus();
                return;
            }

            DialogResult vRESULT;
            vRESULT = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (vRESULT == DialogResult.Yes)
            {
                Application.UseWaitCursor = true;
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                string mSTATUS = "F";
                string mMESSAGE = string.Empty;
                idcSET_PAYMENT_DELETE.ExecuteNonQuery();
                mSTATUS = idcSET_PAYMENT_DELETE.GetCommandParamValue("O_STATUS").ToString();
                mMESSAGE = iString.ISNull(idcSET_PAYMENT_DELETE.GetCommandParamValue("O_MESSAGE"));
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (idcSET_PAYMENT_DELETE.ExcuteError || mSTATUS == "F")
                {
                    if (mMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
                if (mMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }            
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }
    }
}