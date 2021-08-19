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

namespace HRMF0410
{
    public partial class HRMF0410 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #endregion;

        #region ----- Constructor -----

        public HRMF0410()
        {
            InitializeComponent();
        }

        public HRMF0410(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }
            IDA_SALARY_TOTAL.Fill();
        }

        private bool Create_Data()
        {
            bool vReturn_Value = false;

            string vSTATUS = "N";
            string vMESSAGE = string.Empty;

            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return vReturn_Value;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return vReturn_Value;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_SALARY_TOTAL.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_SALARY_TOTAL.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_SALARY_TOTAL.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_SALARY_TOTAL.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return vReturn_Value;
            }

            vReturn_Value = true;
            return vReturn_Value;
        }

        private bool Salary_Total_Closed(string pCLOSED_STATUS)
        {
            bool vReturn_Value = false;

            string vSTATUS = "N";
            string vMESSAGE = string.Empty;

            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return vReturn_Value;
            }
            if (iConv.ISNull(W_PERIOD_NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return vReturn_Value;
            }
            if (pCLOSED_STATUS == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10502"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vReturn_Value;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_SET_SALARY_TOTAL_CLOSED.SetCommandParamValue("P_CLOSED_STATUS", pCLOSED_STATUS);
            IDC_SET_SALARY_TOTAL_CLOSED.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_SALARY_TOTAL_CLOSED.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_SALARY_TOTAL_CLOSED.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_SALARY_TOTAL_CLOSED.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return vReturn_Value;
            }

            vReturn_Value = true;
            return vReturn_Value;
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

        #region ----- Form Event ------
        
        private void HRMF0410_Load(object sender, EventArgs e)
        {
            // Lookup SETTING
            ILD_W_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ILD_W_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            // Standard Date SETTING
            //DateTime dLastYearMonthDay = new DateTime(DateTime.Today.Year, 12, 31);
            //STD_YYYYMM.EditValue = dLastYearMonthDay;
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);

            IDA_SALARY_TOTAL.FillSchema();
        }

        private void BTN_CREATE_DATA_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Create_Data() == false)
            {
                return;
            }

            //다시 조회 
            Search_DB();
        }

        private void BTN_SET_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Salary_Total_Closed("OK") == false)
            {
                return;
            }

            //다시 조회 
            Search_DB();
        }

        private void BTN_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Salary_Total_Closed("CANCEL") == false)
            {
                return;
            }

            //다시 조회 
            Search_DB();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_W_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISDate_Month_Add(DateTime.Today, 5));
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_EMPLYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_W_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
        }
        
        #endregion

    }
}