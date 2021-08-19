using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using Syncfusion.Windows.Forms;
using System;
using System.Data;
using System.Windows.Forms;

namespace HRMF0714
{
    public partial class HRMF0714 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mUSER_CAP = "N";

        #endregion;

        #region ----- Constructor -----

        public HRMF0714()
        {
            InitializeComponent();
        }

        public HRMF0714(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- User Make Methods ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void DefaultDate()
        {
            if (DateTime.Today.Month <= 2)
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(iDate.ISDate_Add(string.Format("{0}-01-01", DateTime.Today.Year), -1));
            }
            else
            {
                W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            }
        }

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
                CB_BatchCreate.Visible = false;
                CB_DONATION_ADJUST_ALL.Visible = false;
                CB_BatchCreate.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_DONATION_ADJUST_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            }
            else
            {
                CB_BatchCreate.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_DONATION_ADJUST_ALL.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_BatchCreate.Visible = true;
                CB_DONATION_ADJUST_ALL.Visible = true;
            }
        }

        private bool Closing_Check()
        {
            if (IGR_PERSON.RowIndex < 0)
            {
                MessageBoxAdv.Show("사원정보가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return false;
            }
            IDC_CLOSING_CHECK_20_P.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CLOSING_CHECK_20_P.GetCommandParamValue("O_MESSAGE"));
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

        private void Init_Grid(string pYear_YYYYMM)
        {
            //그리드//
            if (iDate.ISGetDate(string.Format("{0}-01-01", pYear_YYYYMM.Substring(0, 4))) < iDate.ISGetDate("2017-01-01"))
            {
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("CREDIT_ALL_AMT_2014")].Visible = 1;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT_2014")].Visible = 1;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("PRE_CREDIT_ALL_AMT")].Visible = 1;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT")].Visible = 1;

                BTN_PRE_CREDIT_UPDATE.Visible = true;
            }
            else
            {
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("CREDIT_ALL_AMT_2014")].Visible = 0;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT_2014")].Visible = 0;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("PRE_CREDIT_ALL_AMT")].Visible = 0;
                IGR_FAMILY_AMOUNT.GridAdvExColElement[IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT")].Visible = 0;

                BTN_PRE_CREDIT_UPDATE.Visible = false;
            }
        }
          
        private DateTime GetDateTime()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        private void SEARCH_DB()
        {
            string vMessage = string.Empty;
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_STD_YYYYMM.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }

            Init_Grid(iString.ISNull(W_STD_YYYYMM.EditValue)); 

            IGR_PERSON.LastConfirmChanges();
            IDA_PERSON.OraSelectData.AcceptChanges();
            IDA_PERSON.Refillable = true;

            IGR_FAMILY.LastConfirmChanges();
            IDA_FAMILY.OraSelectData.AcceptChanges();
            IDA_FAMILY.Refillable = true;

            IGR_FAMILY_AMOUNT.LastConfirmChanges();
            IDA_FAMILY_AMOUNT.OraSelectData.AcceptChanges();
            IDA_FAMILY_AMOUNT.Refillable = true;

            IGR_FAMILY_CARD_AMOUNT.LastConfirmChanges();
            IDA_FAMILY_CARD_AMOUNT.OraSelectData.AcceptChanges();
            IDA_FAMILY_CARD_AMOUNT.Refillable = true;

            IDA_FOUNDATION.OraSelectData.AcceptChanges();
            IDA_FOUNDATION.Refillable = true;

            IGR_SAVING_INFO.LastConfirmChanges();
            IDA_SAVING_INFO.OraSelectData.AcceptChanges();
            IDA_SAVING_INFO.Refillable = true;

            IGR_SAVING_INVEST.LastConfirmChanges();
            IDA_SAVING_INVEST.OraSelectData.AcceptChanges();
            IDA_SAVING_INVEST.Refillable = true;

            IGR_DONATION_INFO.LastConfirmChanges();
            IDA_DONATION_INFO.OraSelectData.AcceptChanges();
            IDA_DONATION_INFO.Refillable = true;

            IGR_DONATION_ADJUSTMENT.LastConfirmChanges();
            IDA_DONATION_ADJUSTMENT.OraSelectData.AcceptChanges();
            IDA_DONATION_ADJUSTMENT.Refillable = true;

            IGR_HOUSE_LEASE_INFO_10.LastConfirmChanges();
            IDA_HOUSE_LEASE_INFO_10.OraSelectData.AcceptChanges();
            IDA_HOUSE_LEASE_INFO_10.Refillable = true;

            IGR_HOUSE_LEASE_INFO_20_1.LastConfirmChanges();
            IGR_HOUSE_LEASE_INFO_20_2.LastConfirmChanges();
            IDA_HOUSE_LEASE_INFO_20.OraSelectData.AcceptChanges();
            IDA_HOUSE_LEASE_INFO_20.Refillable = true; 

            try
            {
                string vPERSON_NUM = iString.ISNull(IGR_PERSON.GetCellValue("PERSON_NUM"));
                int vIDX_Col = IGR_PERSON.GetColumnToIndex("PERSON_NUM");

                IDA_PERSON.Fill();
                if (IGR_PERSON.RowCount > 0)
                {
                    for (int vRow = 0; vRow < IGR_PERSON.RowCount; vRow++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_PERSON.GetCellValue(vRow, vIDX_Col)))
                        {
                            IGR_PERSON.CurrentCellActivate(vRow, 0);
                            IGR_PERSON.CurrentCellMoveTo(vRow, 0);
                        }
                    }
                }
                IGR_PERSON.Focus();
            }
            catch (System.Exception ex)
            {
                vMessage = string.Format("Adapter Fill Error\n{0}", ex.Message);
                MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SET_TB_ADJUSTMENT_FOCUS()
        {
            if (TB_ADJUSTMENT.SelectedTab.TabIndex == 1)
            {
                REPRE_NUM.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 2)
            {
                IGR_FAMILY.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 3)
            {
                STOCK_BENE_AMT.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 4)
            {
                ANNU_INSUR_AMT.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 5)
            {
                IGR_SAVING_INFO.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 6)
            {
                IGR_DONATION_INFO.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 7)
            {
                IGR_DONATION_ADJUSTMENT.Focus();
            }
            else if (TB_ADJUSTMENT.SelectedTab.TabIndex == 8)
            {
                IGR_HOUSE_LEASE_INFO_10.Focus();
            }
        }

        private void SetCommon(object pGROUP_CODE, object pENABLED_FLAG_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pENABLED_FLAG_YN);
        }

        private void CREATE_SUPPORT_FAMILY()
        {
            string vMessage = string.Empty;
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_STD_YYYYMM.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            if (CB_BatchCreate.CheckedState == ISUtil.Enum.CheckedState.Unchecked && iString.ISNull(PERSON_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("[개별생성] {0}", isMessageAdapter1.ReturnText("FCM_10028")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Closing_Check() == false)
            {
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            isDataTransaction1.BeginTran();

            if (iString.ISNull(ADJUST_YYYY.EditValue) == String.Empty)
            {
                string mYYYY = iDate.ISYear(iDate.ISGetDate(W_STD_YYYYMM.EditValue));
                IDC_FAMAILY_CREATE.SetCommandParamValue("P_YEAR_YYYY", mYYYY);
            }
            else
            {
                IDC_FAMAILY_CREATE.SetCommandParamValue("P_YEAR_YYYY", ADJUST_YYYY.EditValue);
            }
            if (CB_BatchCreate.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
            {
                IDC_FAMAILY_CREATE.SetCommandParamValue("P_PERSON_ID", PERSON_ID.EditValue);
            }
            else
            {
                IDC_FAMAILY_CREATE.SetCommandParamValue("P_PERSON_ID", System.DBNull.Value);
            }

            IDC_FAMAILY_CREATE.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_FAMAILY_CREATE.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_FAMAILY_CREATE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (IDC_FAMAILY_CREATE.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            isDataTransaction1.Commit();
            isAppInterfaceAdv1.OnAppMessage(mMESSAGE);
        }

        private void Init_Pre_Credit_Update()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("정산년도를 선택하지 않았습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(IGR_PERSON.GetCellValue("PERSON_ID")) == string.Empty)
            {
                MessageBoxAdv.Show("사원을 선택하지 않았습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;

            IDC_INIT_PRE_CREDIT_AMOUNT.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_INIT_PRE_CREDIT_AMOUNT.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_INIT_PRE_CREDIT_AMOUNT.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (IDC_INIT_PRE_CREDIT_AMOUNT.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            isAppInterfaceAdv1.OnAppMessage(string.Empty);
        }

        private void CREATE_DONATION_ADJUSTMENT()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_STD_YYYYMM.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            if (CB_BatchCreate.CheckedState == ISUtil.Enum.CheckedState.Unchecked && iString.ISNull(PERSON_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("[개별생성] {0}", isMessageAdapter1.ReturnText("FCM_10028")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Closing_Check() == false)
            {
                return;
            }

            DialogResult vdlgResult;
            vdlgResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (vdlgResult == DialogResult.No)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = null;
            isDataTransaction1.BeginTran();

            if (iString.ISNull(ADJUST_YYYY.EditValue) == String.Empty)
            {
                string mYYYY = iDate.ISYear(iDate.ISGetDate(W_STD_YYYYMM.EditValue));
                IDC_DONATION_ADJUSTMENT.SetCommandParamValue("P_YEAR_YYYY", mYYYY);
            }
            else
            {
                IDC_DONATION_ADJUSTMENT.SetCommandParamValue("P_YEAR_YYYY", ADJUST_YYYY.EditValue);
            }

            if (CB_BatchCreate.CheckedState == ISUtil.Enum.CheckedState.Unchecked)
            {
                IDC_DONATION_ADJUSTMENT.SetCommandParamValue("P_PERSON_ID", PERSON_ID.EditValue);
            }
            else
            {
                IDC_DONATION_ADJUSTMENT.SetCommandParamValue("P_PERSON_ID", System.DBNull.Value);
            }

            IDC_DONATION_ADJUSTMENT.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_DONATION_ADJUSTMENT.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_DONATION_ADJUSTMENT.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            if (IDC_DONATION_ADJUSTMENT.ExcuteError || mSTATUS == "F")
            {
                isDataTransaction1.RollBack();
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            isDataTransaction1.Commit();
            isAppInterfaceAdv1.OnAppMessage(mMESSAGE);

            MessageBoxAdv.Show("기부금 조정명세서 생성을 완료하였습니다. \r\n연말정산 계산을 하셔야 [기부금 해당연도 공제금액]이 반영됩니다.", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Show_Address_Live()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, LIVE_ZIP_CODE.EditValue, LIVE_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                LIVE_ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                LIVE_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_House_Lease_10()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                                IGR_HOUSE_LEASE_INFO_10.GetCellValue("LEASE_ZIP_CODE"),
                                                                IGR_HOUSE_LEASE_INFO_10.GetCellValue("LEASE_ADDR1"));
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                IGR_HOUSE_LEASE_INFO_10.SetCellValue("LEASE_ZIP_CODE", vEAPF0299.Get_Zip_Code);
                IGR_HOUSE_LEASE_INFO_10.SetCellValue("LEASE_ADDR1", vEAPF0299.Get_Address);
            }
            vEAPF0299.Dispose();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_House_Lease_20()
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                                IGR_HOUSE_LEASE_INFO_20_2.GetCellValue("LEASE_ZIP_CODE"),
                                                                IGR_HOUSE_LEASE_INFO_20_2.GetCellValue("LEASE_ADDR1"));
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                IGR_HOUSE_LEASE_INFO_20_2.SetCellValue("LEASE_ZIP_CODE", vEAPF0299.Get_Zip_Code);
                IGR_HOUSE_LEASE_INFO_20_2.SetCellValue("LEASE_ADDR1", vEAPF0299.Get_Address);
            }
            vEAPF0299.Dispose();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        #endregion;

        #region ----- 주민번호 체크 ------

        private string REPRE_NUM_CHECK(object pREPRE_NUM)
        {
            string isReturnValue = "N".ToString();
            if (iString.ISNull(pREPRE_NUM) == string.Empty)
            {
                return isReturnValue;
            }
            if (iString.ISNull(pREPRE_NUM).Replace("-", "").Length < 13)
            {
                IDC_CHECK_TAX_NUM.SetCommandParamValue("P_TAX_NUM", pREPRE_NUM);
                IDC_CHECK_TAX_NUM.ExecuteNonQuery();

                isReturnValue = IDC_CHECK_TAX_NUM.GetCommandParamValue("O_RETURN_VALUE").ToString();
            }
            else
            {
                IDC_REPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pREPRE_NUM);
                IDC_REPRE_NUM_CHECK.ExecuteNonQuery();

                isReturnValue = IDC_REPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            }
            return isReturnValue;
        }

        #endregion;

        #region ----- Main Button Events -----

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
                    if (Closing_Check() == false)
                    {
                        return;
                    }

                    if (IDA_SAVING_INFO.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 5))
                    {
                        IDA_SAVING_INFO.AddOver();
                        IGR_SAVING_INFO.Focus();
                    }
                    else if (IDA_SAVING_INVEST.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 5))
                    {
                        IDA_SAVING_INVEST.AddOver();
                        IGR_SAVING_INVEST.Focus();
                    }
                    else if (IDA_DONATION_INFO.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 6))
                    {
                        IDA_DONATION_INFO.AddOver();
                        IGR_DONATION_INFO.Focus();
                    }
                    else if (IDA_DONATION_ADJUSTMENT.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 7))
                    {
                        IDA_DONATION_ADJUSTMENT.AddOver();
                        IGR_DONATION_ADJUSTMENT.Focus();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_10.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 8))
                    {
                        IDA_HOUSE_LEASE_INFO_10.AddOver();
                        IGR_HOUSE_LEASE_INFO_10.Focus();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_20.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 8))
                    {
                        IDA_HOUSE_LEASE_INFO_20.AddOver();
                        IGR_HOUSE_LEASE_INFO_20_1.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (Closing_Check() == false)
                    {
                        return;
                    }

                    if (IDA_SAVING_INFO.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 5))
                    {
                        IDA_SAVING_INFO.AddUnder();
                        IGR_SAVING_INFO.Focus();
                    }
                    else if (IDA_SAVING_INVEST.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 5))
                    {
                        IDA_SAVING_INVEST.AddUnder();
                        IGR_SAVING_INVEST.Focus();
                    }
                    else if (IDA_DONATION_INFO.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 6))
                    {
                        IDA_DONATION_INFO.AddUnder();
                        IGR_DONATION_INFO.Focus();
                    }
                    else if (IDA_DONATION_ADJUSTMENT.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 7))
                    {
                        IDA_DONATION_ADJUSTMENT.AddUnder();
                        IGR_DONATION_ADJUSTMENT.Focus();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_10.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 8))
                    {
                        IDA_HOUSE_LEASE_INFO_10.AddUnder();
                        IGR_HOUSE_LEASE_INFO_10.Focus();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_20.IsFocused || (IDA_PERSON.IsFocused && TB_ADJUSTMENT.SelectedTab.TabIndex == 8))
                    {
                        IDA_HOUSE_LEASE_INFO_20.AddUnder();
                        IGR_HOUSE_LEASE_INFO_20_1.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    NAME.Focus();
                    if (Closing_Check() == false)
                    {
                        return;
                    }

                    try
                    {
                        IDA_PERSON.Update();
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_PERSON.IsFocused)
                    {
                        IDA_HOUSE_LEASE_INFO_20.Cancel();
                        IDA_HOUSE_LEASE_INFO_10.Cancel();
                        IDA_DONATION_ADJUSTMENT.Cancel();
                        IDA_DONATION_INFO.Cancel();
                        IDA_SAVING_INFO.Cancel();
                        IDA_SAVING_INVEST.Cancel();
                        IDA_FAMILY_CARD_AMOUNT.Cancel();
                        IDA_FAMILY_AMOUNT.Cancel();
                        IDA_FAMILY.Cancel();
                        IDA_PERSON.Cancel();
                    }
                    else if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.Cancel();
                    }
                    else if (IDA_FAMILY_AMOUNT.IsFocused)
                    {
                        IDA_FAMILY_AMOUNT.Cancel();
                    }
                    else if(IDA_FAMILY_CARD_AMOUNT.IsFocused)
                    {
                        IDA_FAMILY_CARD_AMOUNT.Cancel();
                    }
                    else if (IDA_FOUNDATION.IsFocused)
                    {
                        IDA_FOUNDATION.Cancel();
                    }
                    else if (IDA_SAVING_INFO.IsFocused)
                    {
                        IDA_SAVING_INFO.Cancel();
                    }
                    else if (IDA_SAVING_INVEST.IsFocused)
                    {
                        IDA_SAVING_INVEST.Cancel();
                    }
                    else if (IDA_DONATION_INFO.IsFocused)
                    {
                        IDA_DONATION_INFO.Cancel();
                    }
                    else if (IDA_DONATION_ADJUSTMENT.IsFocused)
                    {
                        IDA_DONATION_ADJUSTMENT.Cancel();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_10.IsFocused)
                    {
                        IDA_HOUSE_LEASE_INFO_10.Cancel();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_20.IsFocused)
                    {
                        IDA_HOUSE_LEASE_INFO_20.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (Closing_Check() == false)
                    {
                        return;
                    }

                    if (IDA_FAMILY.IsFocused)
                    {
                        IDA_FAMILY.Delete();
                    }
                    else if (IDA_FAMILY_AMOUNT.IsFocused)
                    {
                        IDA_FAMILY_AMOUNT.Delete();
                    }
                    else if (IDA_FAMILY_CARD_AMOUNT.IsFocused)
                    {
                        IDA_FAMILY_CARD_AMOUNT.Delete();
                    }
                    else if (IDA_FOUNDATION.IsFocused)
                    {
                        IDA_FOUNDATION.Delete();
                    }
                    else if (IDA_SAVING_INFO.IsFocused)
                    {
                        IDA_SAVING_INFO.Delete();
                    }
                    else if (IDA_SAVING_INVEST.IsFocused)
                    {
                        IDA_SAVING_INVEST.Delete();
                    }
                    else if (IDA_DONATION_INFO.IsFocused)
                    {
                        IDA_DONATION_INFO.Delete();
                    }
                    else if (IDA_DONATION_ADJUSTMENT.IsFocused)
                    {
                        IDA_DONATION_ADJUSTMENT.Delete();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_10.IsFocused)
                    {
                        IDA_HOUSE_LEASE_INFO_10.Delete();
                    }
                    else if (IDA_HOUSE_LEASE_INFO_20.IsFocused)
                    {
                        IDA_HOUSE_LEASE_INFO_20.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                }
            }
        }

        #endregion;

        #region ----- This Form Events -----

        private void HRMF0714_Load(object sender, EventArgs e)
        {
            IDA_PERSON.FillSchema();
            IDA_FAMILY.FillSchema();
            IDA_FAMILY_AMOUNT.FillSchema();
            IDA_FOUNDATION.FillSchema();
            IDA_SAVING_INFO.FillSchema();
            IDA_DONATION_INFO.FillSchema();
            IDA_DONATION_ADJUSTMENT.FillSchema();
            IDA_HOUSE_LEASE_INFO_10.FillSchema();
            IDA_HOUSE_LEASE_INFO_20.FillSchema();
        }

        private void HRMF0714_Shown(object sender, EventArgs e)
        {
            DefaultDate();
            DefaultCorporation();
            CB_BatchCreate.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            PM_DONATION.PromptTextElement[0].Default = "기부금명세서를 작성하셨을 경우 [기부금 조정명세서]를 생성하시기 바랍니다.\r\n해당연도 기부금 공제대상금액이 생성됩니다";
            Init_Grid(iString.ISNull(W_STD_YYYYMM.EditValue));
            User_Cap();
        }

        private void BTN_FAMILY_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            CREATE_SUPPORT_FAMILY();
        }

        private void BTN_DONATION_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            CREATE_DONATION_ADJUSTMENT();
        }

        private void TB_ADJUSTMENT_Click(object sender, EventArgs e)
        {
            SET_TB_ADJUSTMENT_FOCUS();
        }

        private void IGR_FAMILY_AMOUNT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {

        }

        private void IGR_DONATION_ADJUSTMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int mIDX_DOAN_AMT = IGR_DONATION_ADJUSTMENT.GetColumnToIndex("DONA_AMT");
            int mIDX_PRE_DONA_DED_AMT = IGR_DONATION_ADJUSTMENT.GetColumnToIndex("PRE_DONA_DED_AMT");
            decimal mTOTAL_DONA_AMT = 0;
            if (e.ColIndex == mIDX_DOAN_AMT)
            {
                mTOTAL_DONA_AMT = iString.ISDecimaltoZero(e.NewValue) -
                                    iString.ISDecimaltoZero(IGR_DONATION_ADJUSTMENT.GetCellValue("PRE_DONA_DED_AMT"));
                IGR_DONATION_ADJUSTMENT.SetCellValue("TOTAL_DONA_AMT", mTOTAL_DONA_AMT);
            }
            else if (e.ColIndex == mIDX_PRE_DONA_DED_AMT)
            {
                mTOTAL_DONA_AMT = iString.ISDecimaltoZero(IGR_DONATION_ADJUSTMENT.GetCellValue("DONA_AMT")) -
                                    iString.ISDecimaltoZero(e.NewValue);
                IGR_DONATION_ADJUSTMENT.SetCellValue("TOTAL_DONA_AMT", mTOTAL_DONA_AMT);
            }
        }

        private void LIVE_ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Live();
            }
        }

        private void LIVE_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Live();
            }
        }

        private void IGR_FAMILY_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_BASE_YN = IGR_FAMILY.GetColumnToIndex("BASE_YN");
            if (e.ColIndex == vIDX_BASE_YN)
            {
                IDC_INIT_SUPPORT_FAMILY_YN_P.SetCommandParamValue("W_BASE_YN", e.NewValue);
                IDC_INIT_SUPPORT_FAMILY_YN_P.ExecuteNonQuery();

                IGR_FAMILY.SetCellValue("SPOUSE_YN", IDC_INIT_SUPPORT_FAMILY_YN_P.GetCommandParamValue("O_SPOUSE_YN"));
                IGR_FAMILY.SetCellValue("OLD_YN", IDC_INIT_SUPPORT_FAMILY_YN_P.GetCommandParamValue("O_OLD_YN"));
                IGR_FAMILY.SetCellValue("OLD1_YN", IDC_INIT_SUPPORT_FAMILY_YN_P.GetCommandParamValue("O_OLD1_YN"));
                IGR_FAMILY.SetCellValue("CHILD_YN", IDC_INIT_SUPPORT_FAMILY_YN_P.GetCommandParamValue("O_CHILD_YN"));
                IGR_FAMILY.SetCellValue("BIRTH_YN", IDC_INIT_SUPPORT_FAMILY_YN_P.GetCommandParamValue("O_BIRTH_YN"));
            }
            //if (e.ColIndex == vIDX_BASE_YN && iString.ISNull(IGR_FAMILY.GetCellValue("YEAR_RELATION_CODE")) == "3")
            //{
            //    IGR_FAMILY.SetCellValue("SPOUSE_YN", e.NewValue);
            //}
        }

        private void IGR_HOUSE_LEASE_INFO_10_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && IGR_HOUSE_LEASE_INFO_10.ColIndex == IGR_HOUSE_LEASE_INFO_10.GetColumnToIndex("LEASE_ADDR1"))
            {
                Show_Address_House_Lease_10();
            }
        }

        private void IGR_HOUSE_LEASE_INFO_10_CellDoubleClick(object pSender)
        {
            if (IGR_HOUSE_LEASE_INFO_10.ColIndex == IGR_HOUSE_LEASE_INFO_10.GetColumnToIndex("LEASE_ADDR1"))
            {
                Show_Address_House_Lease_10();
            }
        }

        private void IGR_HOUSE_LEASE_INFO_20_2_CellDoubleClick(object pSender)
        {
            if (IGR_HOUSE_LEASE_INFO_20_2.ColIndex == IGR_HOUSE_LEASE_INFO_20_2.GetColumnToIndex("LEASE_ADDR1"))
            {
                Show_Address_House_Lease_20();
            }
        }

        private void IGR_HOUSE_LEASE_INFO_20_2_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && IGR_HOUSE_LEASE_INFO_20_2.ColIndex == IGR_HOUSE_LEASE_INFO_20_2.GetColumnToIndex("LEASE_ADDR1"))
            {
                Show_Address_House_Lease_20();
            }
        }

        private void BTN_PRE_CREDIT_UPDATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Init_Pre_Credit_Update();
        }

        private void BTN_PRE_COPY_ButtonClick(object pSender, EventArgs pEventArgs)
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

            HRMF0714_COPY vHRMF0714_COPY = new HRMF0714_COPY(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                            , W_STD_YYYYMM.EditValue
                                                            , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                            , null, null
                                                            , W_FLOOR_NAME.EditValue, W_FLOOR_ID.EditValue
                                                            , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0714_COPY.ShowDialog();
            vHRMF0714_COPY.Dispose();
        }



        #endregion;

        #region ----- Lookup Event -----

        private void ilaYYYYMM_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISDate_Month_Add(iDate.ISGetDate(), 4));
        }

        private void ilaYYYYMM_SelectedRowData(object pSender)
        {
            User_Cap();
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_EMPLOYE_TYPE", "Y");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("FLOOR", "Y");
        }

        private void ilaEDUCATION_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("EDU_LMT", "Y");
        }

        private void ILA_RESIDENT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("RESIDENT_TYPE", "Y");
        }

        private void ILA_NATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("NATION", "Y");
        }

        private void ILA_NATIONALITY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("NATIONALITY_TYPE", "Y");
        }

        private void ILA_HOUSEHOLD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("HOUSEHOLD_TYPE", "Y");
        }

        private void ilaADDRESS_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildADDRESS.SetLookupParamValue("W_ADDRESS", LIVE_ZIP_CODE.EditValue);
        }

        private void ILA_SAVING_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SAVING_TYPE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_YEAR_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_BANK", "Y");
        }

        private void ILA_DONATION_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("DONATION_TYPE", "Y");
        }

        private void ILA_YEAR_RELATION_5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_RELATION", "Y");
        }

        private void ILA_DONATION_TYPE_6_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("DONATION_TYPE", "Y");
        }

        private void ILA_DISABILITY_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_DISABILITY", "Y");
        }

        private void ILA_HOUSE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("HOUSE_TYPE", "Y");
        }

        private void ILA_HOUSE_TYPE2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("HOUSE_TYPE", "Y");
        }

        private void ILA_SAVING_TYPE_IV_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_SAVING_TYPE_IV.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_INVEST_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_INVEST_TYPE", "Y");
        }

        private void ILA_YEAR_BANK_IV_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("YEAR_BANK", "Y");
        }

        private void ILA_DONATION_CONTENTS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon("DONATION_CONTENTS", "Y");
        }

        #endregion

        #region ----- Adapter Event -----


        private void IDA_FAMILY_AMOUNT_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            int vIDX_AMOUNT_TYPE = IGR_FAMILY_AMOUNT.GetColumnToIndex("AMOUNT_TYPE");
            for (int r = 0; r < IGR_FAMILY_AMOUNT.RowCount; r++)
            {
                if ("99" == iString.ISNull(IGR_FAMILY_AMOUNT.GetCellValue(r, vIDX_AMOUNT_TYPE)))
                {
                    IGR_FAMILY_AMOUNT.RowBackColor_Sum(r);
                    return;
                }
            }
        }

        private void IDA_FAMILY_AMOUNT_FilterCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            int vIDX_AMOUNT_TYPE = IGR_FAMILY_AMOUNT.GetColumnToIndex("AMOUNT_TYPE");
            for (int r = 0; r < IGR_FAMILY_AMOUNT.RowCount; r++)
            {
                if ("99" == iString.ISNull(IGR_FAMILY_AMOUNT.GetCellValue(r, vIDX_AMOUNT_TYPE)))
                {
                    IGR_FAMILY_AMOUNT.RowBackColor_Sum(r);
                    return;
                }
            }
        }

        private void idaPERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=[Person No]"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RESIDENT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[거주구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NATION_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show("[국가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NATIONALITY_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[내외국인 구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["HOUSEHOLD_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[세대주 구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LIVE_ADDR1"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주소]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_FAMILY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["YEAR_RELATION_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[관계]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FAMILY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DISABILITY_YN"]) == "Y" && iString.ISNull(e.Row["DISABILITY_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[장애인]을 선택했을 경우 장애인구분을 입력해야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DISABILITY_YN"]) == "N" && iString.ISNull(e.Row["DISABILITY_CODE"]) != string.Empty)
            {
                MessageBoxAdv.Show("[장애인구분]을 선택했을 경우 장애인 여부를 선택해야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            //인적공제 처리 검증//
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("W_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("W_REPRE_NUM", e.Row["REPRE_NUM"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("W_YEAR_YYYY", e.Row["YEAR_YYYY"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_YEAR_RELATION_CODE", e.Row["YEAR_RELATION_CODE"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_BASE_YN", e.Row["BASE_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_BASE_LIVING_YN", e.Row["BASE_LIVING_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_SPOUSE_YN", e.Row["SPOUSE_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_OLD_YN", e.Row["OLD_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_OLD1_YN", e.Row["OLD1_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_WOMAN_YN", e.Row["WOMAN_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_SINGLE_PARENT_DED_YN", e.Row["SINGLE_PARENT_DED_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_CHILD_YN", e.Row["CHILD_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_BIRTH_YN", e.Row["BIRTH_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_DISABILITY_YN", e.Row["DISABILITY_YN"]);
            IDC_CHECK_SUPPORT_FAMILY.SetCommandParamValue("P_DISABILITY_CODE", e.Row["DISABILITY_CODE"]);
            IDC_CHECK_SUPPORT_FAMILY.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CHECK_SUPPORT_FAMILY.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CHECK_SUPPORT_FAMILY.GetCommandParamValue("O_MESSAGE"));
            if (IDC_CHECK_SUPPORT_FAMILY.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                e.Cancel = true;
                return;
            }
        }

        private void IDA_FAMILY_AMOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //검증//
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("W_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("W_REPRE_NUM", e.Row["REPRE_NUM"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("W_YEAR_YYYY", e.Row["YEAR_YYYY"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("W_AMOUNT_TYPE", e.Row["AMOUNT_TYPE"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("W_YEAR_RELATION_CODE", e.Row["YEAR_RELATION_CODE"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_MEDIC_INSUR_AMT", e.Row["MEDIC_INSUR_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_HIRE_INSUR_AMT", e.Row["MEDIC_INSUR_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_INSURE_AMT", e.Row["INSURE_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_DISABILITY_INSURE_AMT", e.Row["DISABILITY_INSURE_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_MEDICAL_AMT", e.Row["MEDICAL_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_MEDICAL_NANIM_AMT", e.Row["MEDICAL_NANIM_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_EDUCATION_AMT", e.Row["EDU_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_CREDIT_AMT", e.Row["CREDIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_CHECK_CREDIT_AMT", e.Row["CHECK_CREDIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_CASH_AMT", e.Row["CASH_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_ACADE_GIRO_AMT", e.Row["ACADE_GIRO_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_TRAD_MARKET_AMT", e.Row["CREDIT_TRAD_MARKET_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_PUBLIC_TRANSIT_AMT", e.Row["CREDIT_PUBLIC_TRANSIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_CREDIT_ALL_AMT_2013", e.Row["CREDIT_ALL_AMT_2013"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_ADD_CREDIT_AMT_2013", e.Row["ADD_CREDIT_AMT_2013"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_CREDIT_ALL_AMT_2014", e.Row["CREDIT_ALL_AMT_2014"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_ADD_CREDIT_AMT_2014", e.Row["ADD_CREDIT_AMT_2014"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_PRE_CREDIT_ALL_AMT", e.Row["PRE_CREDIT_ALL_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_PRE_ADD_CREDIT_AMT", e.Row["PRE_ADD_CREDIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_PRE_SEC_CREDIT_AMT", e.Row["PRE_SEC_CREDIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.SetCommandParamValue("P_ADD_CREDIT_AMT", e.Row["ADD_CREDIT_AMT"]);
            IDC_CHECK_SUPPORT_FAMILY_AMT_P.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CHECK_SUPPORT_FAMILY_AMT_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CHECK_SUPPORT_FAMILY_AMT_P.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(string.Format("Check Data :: {0}", vMESSAGE), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }


            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["YEAR_RELATION_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[관계]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FAMILY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["EDUCATION_TYPE"]) == string.Empty && iString.ISDecimaltoZero(e.Row["EDU_AMT"]) != 0)
            //{
            //    MessageBoxAdv.Show("[교육비 구분]을 선택하지 않고 교육비를 입력했습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["EDUCATION_TYPE"]) != string.Empty && iString.ISDecimaltoZero(e.Row["EDU_AMT"]) == 0)
            //{
            //    MessageBoxAdv.Show("[교육비 구분]을 선택하지 하고 교육비를 입력하지 않았습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISDecimaltoZero(e.Row["EDUCATION_AMOUNT_LMT"]) < iString.ISDecimaltoZero(e.Row["EDU_AMT"]))
            //{
            //    MessageBoxAdv.Show("[교육비 한도]를 초과했습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        private void IDA_SAVING_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SAVING_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[연금저축구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["BANK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[금융기관]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[계좌번호]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SAVING_TYPE"]) == "41" && iString.ISNumtoZero(e.Row["SAVING_COUNT"], 0) == 0)
            {
                MessageBoxAdv.Show("[장기주식형저축소득공제]은 [납입연차]를 입력해야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SAVING_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show("[불입금액]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_SAVING_INVEST_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SAVING_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[연금저축구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["INVEST_YYYY"]) == string.Empty)
            {
                MessageBoxAdv.Show("[투자년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["INVEST_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[투자구분]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["INVEST_NAME"]) == string.Empty && iString.ISNull(e.Row["BANK_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[조합명 또는 투자신탁명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ACCOUNT_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[계좌번호]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SAVING_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show("[불입금액]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_DONATION_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //검증//
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_YEAR_YYYY", ADJUST_YYYY.EditValue);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONA_TYPE", e.Row["DONA_TYPE"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONATION_CONTENTS", e.Row["DONATION_CONTENTS"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_FAMILY_NAME", e.Row["FAMILY_NAME"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_REPRE_NUM", e.Row["REPRE_NUM"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_RELATION_CODE", e.Row["RELATION_CODE"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_CORP_NAME", e.Row["CORP_NAME"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_CORP_TAX_REG_NO", e.Row["CORP_TAX_REG_NO"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONA_DATE", e.Row["DONA_DATE"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_SUB_DESCRIPTION", e.Row["SUB_DESCRIPTION"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONA_COUNT", e.Row["DONA_COUNT"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONA_AMT", e.Row["DONA_AMT"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONATION_SUBSIDY_AMT", e.Row["DONATION_SUBSIDY_AMT"]);
            IDC_CHECK_DONATION_INFO.SetCommandParamValue("P_DONA_BILL_NUM", e.Row["DONA_BILL_NUM"]);
            IDC_CHECK_DONATION_INFO.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CHECK_DONATION_INFO.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CHECK_DONATION_INFO.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(string.Format("Check Data :: {0}", vMESSAGE), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }


            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부금 유형]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FAMILY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부자 성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부자 주민번호]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RELATION_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부자 관계]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORP_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부처 상호]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORP_TAX_REG_NO"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부처 사업자번호]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_COUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부 건수]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_AMT"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부 금액]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_DONATION_ADJUSTMENT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부금 유형]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_YYYY"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부 년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(ADJUST_YYYY.EditValue) < iString.ISNumtoZero(e.Row["DONA_YYYY"]))
            {
                MessageBoxAdv.Show("[기부 년도]가 [정산년도] 이후 일수는 없습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONA_AMT"]) == string.Empty)
            {
                MessageBoxAdv.Show("[기부 금액]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_FAMILY_AMOUNT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            if (iString.ISNull(pBindingManager.DataRow["AMOUNT_TYPE"]) == "99")
            {
                //합계 이므로 그리드 수정 안되게 제어 
                //index = 4 : 기타보험 ~ index = 19 : 기부금(종교)까지 
                for (int c = 4; c < 15; c++)
                {
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[c].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[c].Updatable = 0;
                }
            }
            else
            {
                //합계 이므로 그리드 수정 안되게 제어 
                //index = 4 : 기타보험 ~ index = 19 : 기부금(종교)까지 
                for (int c = 4; c < 27; c++)
                {
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[c].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[c].Updatable = 1;
                }

                int vIDX_CASH_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("CASH_AMT");
                int vIDX_CASH_TRAD_MARKET_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("CASH_TRAD_MARKET_AMT");
                int vIDX_CASH_PUBLIC_TRANSIT_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("CASH_PUBLIC_TRANSIT_AMT");
                int vIDX_CASH_BOOK_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("CASH_BOOK_AMT");

                int vIDX_CREDIT_ALL_AMT_2013 = IGR_FAMILY_AMOUNT.GetColumnToIndex("CREDIT_ALL_AMT_2013");
                int vIDX_ADD_CREDIT_AMT_2013 = IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT_2013");
                int vIDX_CREDIT_ALL_AMT_2014 = IGR_FAMILY_AMOUNT.GetColumnToIndex("CREDIT_ALL_AMT_2014");
                int vIDX_ADD_CREDIT_AMT_2014 = IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT_2014");
                int vIDX_PRE_CREDIT_ALL_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("PRE_CREDIT_ALL_AMT");
                int vIDX_PRE_ADD_CREDIT_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("PRE_ADD_CREDIT_AMT");
                int vIDX_PRE_SEC_CREDIT_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("PRE_SEC_CREDIT_AMT");
                int vIDX_ADD_CREDIT_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("ADD_CREDIT_AMT");
                int vIDX_CREDIT_ALL_AMT = IGR_FAMILY_AMOUNT.GetColumnToIndex("CREDIT_ALL_AMT");

                if (iString.ISNull(pBindingManager.DataRow["AMOUNT_TYPE"]) == "2")
                {
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_TRAD_MARKET_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_TRAD_MARKET_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_PUBLIC_TRANSIT_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_PUBLIC_TRANSIT_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2013].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2013].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2013].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2013].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2014].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2014].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2014].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2014].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_CREDIT_ALL_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_CREDIT_ALL_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_ADD_CREDIT_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_ADD_CREDIT_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_SEC_CREDIT_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_SEC_CREDIT_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT].Updatable = 0;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT].Updatable = 0;
                }
                else
                {
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_TRAD_MARKET_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_TRAD_MARKET_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_PUBLIC_TRANSIT_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_PUBLIC_TRANSIT_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2013].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2013].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2013].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2013].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2014].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT_2014].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2014].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT_2014].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_CREDIT_ALL_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_CREDIT_ALL_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_ADD_CREDIT_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_ADD_CREDIT_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_SEC_CREDIT_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_PRE_SEC_CREDIT_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT].Insertable = 1;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_ADD_CREDIT_AMT].Updatable = 1;

                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT].Insertable = 0;
                    IGR_FAMILY_AMOUNT.GridAdvExColElement[vIDX_CREDIT_ALL_AMT].Updatable = 0;
                }
            }
        }

        private void IDA_FAMILY_CARD_AMOUNT_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            int vIDX_AMOUNT_TYPE = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("AMOUNT_TYPE");
            for (int r = 0; r < IGR_FAMILY_CARD_AMOUNT.RowCount; r++)
            {
                if ("99" == iString.ISNull(IGR_FAMILY_CARD_AMOUNT.GetCellValue(r, vIDX_AMOUNT_TYPE)))
                {
                    IGR_FAMILY_CARD_AMOUNT.RowBackColor_Sum(r);
                    return;
                }
            }
        }

        private void IDA_FAMILY_CARD_AMOUNT_FilterCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            
            int vIDX_AMOUNT_TYPE = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("AMOUNT_TYPE");
            for (int r = 0; r < IGR_FAMILY_CARD_AMOUNT.RowCount; r++)
            {
                if ("99" == iString.ISNull(IGR_FAMILY_CARD_AMOUNT.GetCellValue(r, vIDX_AMOUNT_TYPE)))
                {
                    IGR_FAMILY_CARD_AMOUNT.RowBackColor_Sum(r);
                    return;
                }
            }
        }

        private void IDA_FAMILY_CARD_AMOUNT_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            if (iString.ISNull(pBindingManager.DataRow["AMOUNT_TYPE"]) == "99")
            {
                //합계 이므로 그리드 수정 안되게 제어 
                //index = 4 : 기타보험 ~ index = 19 : 기부금(종교)까지 
                for (int c = 4; c < 45; c++)
                {
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[c].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[c].Updatable = 0;
                }
            }
            else
            {
                //합계 이므로 그리드 수정 안되게 제어 
                //index = 4 : 기타보험 ~ index = 19 : 기부금(종교)까지 
                for (int c = 4; c < 45; c++)
                {
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[c].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[c].Updatable = 1;
                }

                int vIDX_TRADE_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRADE_MAR_AMT");
                int vIDX_TRANS_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRANS_MAR_AMT");

                int vIDX_TRADE_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRADE_APR_AMT");
                int vIDX_TRANS_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRANS_APR_AMT");

                int vIDX_TRADE_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRADE_ETC_AMT");
                int vIDX_TRANS_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("TRANS_ETC_AMT");

                int vIDX_CASH_NOR_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_NOR_MAR_AMT");
                int vIDX_CASH_BOOK_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_BOOK_MAR_AMT");
                int vIDX_CASH_TRADE_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRADE_MAR_AMT");
                int vIDX_CASH_TRANS_MAR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRANS_MAR_AMT");

                int vIDX_CASH_NOR_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_NOR_APR_AMT");
                int vIDX_CASH_BOOK_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_BOOK_APR_AMT");
                int vIDX_CASH_TRADE_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRADE_APR_AMT");
                int vIDX_CASH_TRANS_APR_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRANS_APR_AMT");
                 
                int vIDX_CASH_NOR_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_NOR_ETC_AMT");
                int vIDX_CASH_BOOK_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_BOOK_ETC_AMT");
                int vIDX_CASH_TRADE_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRADE_ETC_AMT");
                int vIDX_CASH_TRANS_ETC_AMT = IGR_FAMILY_CARD_AMOUNT.GetColumnToIndex("CASH_TRANS_ETC_AMT");

                //대중교통, 전통시장 합계//
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_MAR_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_MAR_AMT].Updatable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_MAR_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_MAR_AMT].Updatable = 0;

                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_APR_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_APR_AMT].Updatable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_APR_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_APR_AMT].Updatable = 0;

                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_ETC_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRADE_ETC_AMT].Updatable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_ETC_AMT].Insertable = 0;
                IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_TRANS_ETC_AMT].Updatable = 0;

                if (iString.ISNull(pBindingManager.DataRow["AMOUNT_TYPE"]) == "2")
                {
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_MAR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_MAR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_MAR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_MAR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_MAR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_MAR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_MAR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_MAR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_APR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_APR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_APR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_APR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_APR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_APR_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_APR_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_APR_AMT].Updatable = 0; 

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_ETC_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_ETC_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_ETC_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_ETC_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_ETC_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_ETC_AMT].Updatable = 0;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_ETC_AMT].Insertable = 0;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_ETC_AMT].Updatable = 0; 
                }
                else
                {
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_MAR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_MAR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_MAR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_MAR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_MAR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_MAR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_MAR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_MAR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_APR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_APR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_APR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_APR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_APR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_APR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_APR_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_APR_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_ETC_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_NOR_ETC_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_ETC_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_BOOK_ETC_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_ETC_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRADE_ETC_AMT].Updatable = 1;

                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_ETC_AMT].Insertable = 1;
                    IGR_FAMILY_CARD_AMOUNT.GridAdvExColElement[vIDX_CASH_TRANS_ETC_AMT].Updatable = 1;
                }
            }
        }

        private void IDA_FAMILY_CARD_AMOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //검증//
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("W_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("W_REPRE_NUM", e.Row["REPRE_NUM"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("W_YEAR_YYYY", e.Row["YEAR_YYYY"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("W_AMOUNT_TYPE", e.Row["AMOUNT_TYPE"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("W_YEAR_RELATION_CODE", e.Row["YEAR_RELATION_CODE"]);

            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_NOR_MAR_AMT", e.Row["CREDIT_NOR_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRADE_MAR_AMT", e.Row["CREDIT_TRADE_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRANS_MAR_AMT", e.Row["CREDIT_TRANS_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_BOOK_MAR_AMT", e.Row["CREDIT_BOOK_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_NOR_APR_AMT", e.Row["CREDIT_NOR_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRADE_APR_AMT", e.Row["CREDIT_TRADE_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRANS_APR_AMT", e.Row["CREDIT_TRANS_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_BOOK_APR_AMT", e.Row["CREDIT_BOOK_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_NOR_ETC_AMT", e.Row["CREDIT_NOR_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRADE_ETC_AMT", e.Row["CREDIT_TRADE_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRANS_ETC_AMT", e.Row["CREDIT_TRANS_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_BOOK_ETC_AMT", e.Row["CREDIT_BOOK_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_NOR_MAR_AMT", e.Row["CHECK_NOR_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRADE_MAR_AMT", e.Row["CHECK_TRADE_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRANS_MAR_AMT", e.Row["CHECK_TRANS_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_BOOK_MAR_AMT", e.Row["CHECK_BOOK_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_NOR_APR_AMT", e.Row["CHECK_NOR_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRADE_APR_AMT", e.Row["CHECK_TRADE_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRANS_APR_AMT", e.Row["CHECK_TRANS_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_BOOK_APR_AMT", e.Row["CHECK_BOOK_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_NOR_ETC_AMT", e.Row["CHECK_NOR_ETC_AMT"]); 
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRADE_ETC_AMT", e.Row["CHECK_TRADE_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CHECK_TRANS_ETC_AMT", e.Row["CHECK_TRANS_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CREDIT_TRADE_APR_AMT", e.Row["CHECK_BOOK_ETC_AMT"]); 
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_NOR_MAR_AMT", e.Row["CASH_NOR_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRADE_MAR_AMT", e.Row["CASH_TRADE_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRANS_MAR_AMT", e.Row["CASH_TRANS_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_BOOK_MAR_AMT", e.Row["CASH_BOOK_MAR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_NOR_ETC_AMT", e.Row["CASH_NOR_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRADE_APR_AMT", e.Row["CASH_TRADE_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRANS_APR_AMT", e.Row["CASH_TRANS_APR_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_BOOK_APR_AMT", e.Row["CASH_BOOK_APR_AMT"]); 
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRADE_ETC_AMT", e.Row["CASH_TRADE_ETC_AMT"]);
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_TRANS_ETC_AMT", e.Row["CASH_TRANS_ETC_AMT"]); 
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.SetCommandParamValue("P_CASH_BOOK_ETC_AMT", e.Row["CASH_BOOK_ETC_AMT"]);  
            IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_CHECK_SUPP_FAMILY_CARD_AMT_P.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(string.Format("Check Data :: {0}", vMESSAGE), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }


            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["YEAR_RELATION_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show("[관계]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["FAMILY_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_FOUNDATION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            int vCount = 0;
            //김수신K 요청 : 12년 이전, 이후 동시 입력 가능 요청 
            //if (iString.ISDecimaltoZero(e.Row["LONG_HOUSE_INTER_AMT_1"], 0) != 0)
            //{
            //    vCount = vCount + 1;
            //}
            //if (iString.ISDecimaltoZero(e.Row["LONG_HOUSE_INTER_AMT_2"], 0) != 0)
            //{
            //    vCount = vCount + 1;
            //}
            //if (iString.ISDecimaltoZero(e.Row["LONG_HOUSE_INTER_AMT_3_FIX"], 0) != 0)
            //{
            //    vCount = vCount + 1;
            //}
            //if (iString.ISDecimaltoZero(e.Row["LONG_HOUSE_INTER_AMT_3_ETC"], 0) != 0)
            //{
            //    vCount = vCount + 1;
            //}            
            if (vCount > 1)
            {
                MessageBoxAdv.Show("1,000만원 한도금액, 1,500만원 한도금액 또는 500만원 한도금액중 하나만 입력 가능합니다.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_HOUSE_LEASE_INFO_10_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LESSOR_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대인 성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LESSOR_REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["LESSOR_REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호(사업자번호)가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_ADDR1"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약서 상 주소지]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_TERM_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_TERM_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["LEASE_TERM_TO"]) < iDate.ISGetDate(e.Row["LEASE_TERM_FR"]))
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["MONTLY_LEASE_AMT"], 0) == 0)
            {
                MessageBoxAdv.Show("[월세액]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_HOUSE_LEASE_INFO_20_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(ADJUST_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("[정산년도]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LESSOR_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대인 성명]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LESSOR_REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["LESSOR_REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호(사업자번호)]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_ADDR1"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약서 상 주소지]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_TERM_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LEASE_TERM_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["LEASE_TERM_TO"]) < iDate.ISGetDate(e.Row["LEASE_TERM_FR"]))
            {
                MessageBoxAdv.Show("[임대차계약 기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["DEPOSIT_AMT"], 0) == 0)
            {
                MessageBoxAdv.Show("[전세보증금]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["LOANER_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show("[대주(貸主)]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LOANER_REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show("[주민번호가]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (REPRE_NUM_CHECK(e.Row["LOANER_REPRE_NUM"]) == "N".ToString())
            {
                MessageBoxAdv.Show("[주민번호(사업자번호)]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LOAN_TERM_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show("[금전소비대차계약기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["LOAN_TERM_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show("[금전소비대차계약기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["LOAN_TERM_TO"]) < iDate.ISGetDate(e.Row["LOAN_TERM_FR"]))
            {
                MessageBoxAdv.Show("[금전소비대차계약기간]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["LOAN_INTEREST_RATE"], 0) == 0)
            {
                MessageBoxAdv.Show("[차입금 이자율]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["LOAN_AMT"], 0) == 0)
            {
                MessageBoxAdv.Show("[원리금]이 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["LOAN_INTEREST_RATE"], 0) == 0)
            {
                MessageBoxAdv.Show("[이자]가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}