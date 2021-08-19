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

namespace HRMF0713
{
    public partial class HRMF0713 : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mUSER_CAP = "N";

        #endregion;

        #region ----- Constructor -----

        public HRMF0713()
        {
            InitializeComponent();
        }

        public HRMF0713(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void SearchDB()
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            string vPERSON_NUM = iString.ISNull(IGR_PERSON.GetCellValue("PERSON_NUM"));
            int vIDX_PERSON_NUM = IGR_PERSON.GetColumnToIndex("PERSON_NUM");
            int vIDX_NAME = IGR_PERSON.GetColumnToIndex("NAME");
            try
            {
                IDA_PERSON.Fill();
            }
            catch (System.Exception ex)
            {
                MessageBoxAdv.Show(string.Format("Adapter Fill Error\n{0}", ex.Message), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (IGR_PERSON.RowCount > 0)
            {
                for (int vRow = 0; vRow < IGR_PERSON.RowCount; vRow++)
                {
                    if (vPERSON_NUM == iString.ISNull(IGR_PERSON.GetCellValue(vRow, vIDX_PERSON_NUM)))
                    {
                        IGR_PERSON.CurrentCellMoveTo(vRow, vIDX_NAME);
                    }
                }
            }
            IGR_PERSON.Focus();
        }

        //// Person Info 그리드를 선택 시, Person ID 및 Year 정보를 체크한 후 idaFOUNDATION에 정보를 출력해주는 함수
        //private void SearchDB_YearAdjustment()
        //{
        //    if (iString.ISNull(igrPERSON_INFO.GetCellValue("PERSON_ID")) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        STANDARD_DATE_0.Focus();
        //        return;
        //    }
        //    YEAR_YYYY_0.EditValue = STANDARD_DATE_0.DateTimeValue.Year.ToString();
        //    if (iString.ISNull(YEAR_YYYY_0.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        STANDARD_DATE_0.Focus();
        //        return;
        //    }
        //    idaYEAR_ADJUSTMENT.Fill();
        //}

        private void User_Cap()
        {
            object vSTD_Date = iDate.ISMonth_Last(iDate.ISGetDate(W_STD_YYYYMM.EditValue));
            if (iDate.ISDate(vSTD_Date) == false)
            {
                vSTD_Date = iDate.ISMonth_Last(iDate.ISGetDate(DateTime.Today));
            }
            IDC_USER_CAP_YEAR_ADJUST.SetCommandParamValue("W_START_DATE", vSTD_Date);
            IDC_USER_CAP_YEAR_ADJUST.SetCommandParamValue("W_END_DATE", vSTD_Date);
            IDC_USER_CAP_YEAR_ADJUST.ExecuteNonQuery();
            mUSER_CAP = iString.ISNull(IDC_USER_CAP_YEAR_ADJUST.GetCommandParamValue("O_CAP_LEVEL"));
            if (mUSER_CAP != "C")
            {
                BTN_CLOSED_CANCEL.Visible = false;
                BTN_CLOSED_OK.Visible = false; 
            }
            else
            { 
                BTN_CLOSED_CANCEL.Visible = true;
                BTN_CLOSED_OK.Visible = true;
            }
        }


        private bool Closing_Check()
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return false;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return false;
            }
             
            IDC_CLOSING_CHECK_P.SetCommandParamValue("W_CLOSING_YYYY", iDate.ISYear(iDate.ISGetDate(W_STD_YYYYMM.EditValue)));
            IDC_CLOSING_CHECK_P.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CLOSING_CHECK_P.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS != "S")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }
            return true;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_YEAR_ADJUSTMENT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_YEAR_ADJUSTMENT.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_YEAR_ADJUSTMENT.IsFocused)
                    {
                        IDA_YEAR_ADJUSTMENT.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0713_Load(object sender, EventArgs e)
        {

        }

        private void HRMF0713_Shown(object sender, EventArgs e)
        {
            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            V_RB_CLOSED_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_CLOSED_FLAG.EditValue = V_RB_CLOSED_NO.RadioCheckedString;
             
            // Standard Date SETTING
            // Standard Date SETTING
            //if (DateTime.Today.Month <= 2)
            //{
            //    DateTime dLastYearMonthDay = new DateTime(DateTime.Today.AddYears(-1).Year, 12, 31);
            //    STANDARD_DATE_0.EditValue = dLastYearMonthDay;
            //}
            //else
            //{
            //    DateTime dLastYearMonthDay = new DateTime(DateTime.Today.Year, 12, 31);
            //    STANDARD_DATE_0.EditValue = dLastYearMonthDay;
            //}
            if (DateTime.Today.Month <= 2)
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(iDate.ISDate_Add(string.Format("{0}-01-01", DateTime.Today.Year), -1));
            }
            else
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            }
            YEAR_YYYY.EditValue = W_STD_YYYYMM.DateTimeValue.Year.ToString();

            User_Cap();

            W_PERSON_NAME.Focus();
        }

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            if (V_RB_ALL.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_FLAG.EditValue = V_RB_ALL.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_NO_Click(object sender, EventArgs e)
        {
            if(V_RB_CLOSED_NO.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_FLAG.EditValue = V_RB_CLOSED_NO.RadioCheckedString;
            }
        }

        private void V_RB_CLOSED_YES_Click(object sender, EventArgs e)
        {
            if(V_RB_CLOSED_YES.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_CLOSED_FLAG.EditValue = V_RB_CLOSED_YES.RadioCheckedString;
            }
        }

        private void btnCalculation_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            if(Closing_Check() == false)
            {
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDC_YEAR_ADJUST_SET.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_YEAR_ADJUST_SET.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_YEAR_ADJUST_SET.GetCommandParamValue("O_MESSAGE")); 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_YEAR_ADJUST_SET.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMESSAGE);

            SearchDB();
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            HRMF0713_CLOSED vHRMF0713_CLOSED = new HRMF0713_CLOSED(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , W_STD_YYYYMM.EditValue
                                                                , "N", "Not Closed"
                                                                , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue
                                                                , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0713_CLOSED.ShowDialog();
            vHRMF0713_CLOSED.Dispose();
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            HRMF0713_CLOSED vHRMF0713_CLOSED = new HRMF0713_CLOSED(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , W_STD_YYYYMM.EditValue
                                                                , "Y", "Closed"
                                                                , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue
                                                                , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0713_CLOSED.ShowDialog();
            vHRMF0713_CLOSED.Dispose();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 3)));
        }

        private void ilaYYYYMM_SelectedRowData(object pSender)
        {
            User_Cap();
        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}