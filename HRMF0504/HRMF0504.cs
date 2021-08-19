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

namespace HRMF0504
{
    public partial class HRMF0504 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----
        public HRMF0504(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            if (iString.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }
        #endregion;

        #region ----- Private Methods -----

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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            igbCORP_GROUP_0.BringToFront();
            CORP_NAME_0.BringToFront();
            igbCORP_GROUP_0.Visible = false;

            if (iString.ISNull(CORP_TYPE.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();                

                irb_ALL_0.RadioButtonValue = "A";
                CORP_TYPE.EditValue = "A";
            }
            else if (iString.ISNull(CORP_TYPE.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            
            
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null&& CORP_TYPE.EditValue.ToString() != "4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            string vPERSON_NUM = iString.ISNull(igrPERSON.GetCellValue("PERSON_NUM"));
            idaMONTH_PAYMENT.Fill();
            igrPERSON.Focus();
            if (igrPERSON.RowCount > 0)
            {
                int vIDX_PERSON_NUM = igrPERSON.GetColumnToIndex("PERSON_NUM");
                for (int i = 0; i < igrPERSON.RowCount; i++)
                {
                    if (vPERSON_NUM == iString.ISNull(igrPERSON.GetCellValue(i, vIDX_PERSON_NUM)))
                    {
                        igrPERSON.CurrentCellMoveTo(i, igrPERSON.GetColumnToIndex("NAME"));
                        return;
                    }
                }
            }
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_Flag_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_Flag_YN);
        }

        private void Insert_Month_Allowance()
        {
            igrMONTH_ALLOWANCE.SetCellValue("PERSON_ID", PERSON_ID.EditValue);
            igrMONTH_ALLOWANCE.SetCellValue("PAY_YYYYMM", igrPERSON.GetCellValue("PAY_YYYYMM"));
            igrMONTH_ALLOWANCE.SetCellValue("WAGE_TYPE", igrPERSON.GetCellValue("WAGE_TYPE"));
            igrMONTH_ALLOWANCE.SetCellValue("CORP_ID", igrPERSON.GetCellValue("CORP_ID"));
            igrMONTH_ALLOWANCE.SetCellValue("MONTH_PAYMENT_ID", MONTH_PAYMENT_ID.EditValue);
            igrMONTH_ALLOWANCE.SetCellValue("CREATED_FLAG", "C");
        }

        private void Insert_Month_Deduction()
        {
            igrMONTH_DEDUCTION.SetCellValue("PERSON_ID", PERSON_ID.EditValue);
            igrMONTH_DEDUCTION.SetCellValue("PAY_YYYYMM", igrPERSON.GetCellValue("PAY_YYYYMM"));
            igrMONTH_DEDUCTION.SetCellValue("WAGE_TYPE", igrPERSON.GetCellValue("WAGE_TYPE"));
            igrMONTH_DEDUCTION.SetCellValue("CORP_ID", igrPERSON.GetCellValue("CORP_ID"));
            igrMONTH_DEDUCTION.SetCellValue("MONTH_PAYMENT_ID", MONTH_PAYMENT_ID.EditValue);
            igrMONTH_DEDUCTION.SetCellValue("CREATED_FLAG", "C");
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----

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
                    if (idaMONTH_ALLOWANCE.IsFocused)
                    {
                        idaMONTH_ALLOWANCE.AddOver();
                        Insert_Month_Allowance();
                    }
                    else if (idaMONTH_DEDUCTION.IsFocused)
                    {
                        idaMONTH_DEDUCTION.AddOver();
                        Insert_Month_Deduction();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaMONTH_ALLOWANCE.IsFocused)
                    {
                        idaMONTH_ALLOWANCE.AddUnder();
                        Insert_Month_Allowance();
                    }
                    else if (idaMONTH_DEDUCTION.IsFocused)
                    {
                        idaMONTH_DEDUCTION.AddUnder();
                        Insert_Month_Deduction();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaMONTH_PAYMENT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaMONTH_PAYMENT.IsFocused)
                    {
                        idaMONTH_PAYMENT.Cancel();
                    }
                    else if (idaMONTH_ALLOWANCE.IsFocused)
                    {
                        idaMONTH_ALLOWANCE.Cancel();
                    }
                    else if (idaMONTH_DEDUCTION.IsFocused)
                    {
                        idaMONTH_DEDUCTION.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaMONTH_PAYMENT.IsFocused)
                    {
                        idaMONTH_PAYMENT.Delete();
                    }
                    else if (idaMONTH_ALLOWANCE.IsFocused)
                    {
                        idaMONTH_ALLOWANCE.Delete();
                    }
                    else if (idaMONTH_DEDUCTION.IsFocused)
                    {
                        idaMONTH_DEDUCTION.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0504_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0504_Shown(object sender, EventArgs e)
        {
            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISMonth_Last(DateTime.Today);
            JOB_CATEGORY_NAME.BringToFront();

            DefaultCorporation();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]          

            // LEAVE CLOSE TYPE SETTING
            ildCLOSE_FLAG_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ildCLOSE_FLAG_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            CLOSED_FLAG_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            CLOSED_FLAG_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            System.Windows.Forms.Cursor.Current = Cursors.Default;

            idaMONTH_PAYMENT.FillSchema();
        }

        private void SET_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0504_SET = new HRMF0504_SET(isAppInterfaceAdv1.AppInterface, "CAL"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , PAY_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , DEPT_ID_0.EditValue, DEPT_CODE_0.EditValue, DEPT_NAME_0.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue,PERSON_NAME_0.EditValue
                                                , CORP_TYPE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0504_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0504_SET.ShowDialog();
            vHRMF0504_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtSET_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0504_SET = new HRMF0504_SET(isAppInterfaceAdv1.AppInterface, "CLOSE"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , PAY_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , DEPT_ID_0.EditValue, DEPT_CODE_0.EditValue, DEPT_NAME_0.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue
                                                , CORP_TYPE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0504_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0504_SET.ShowDialog();
            vHRMF0504_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtSET_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(PAY_YYYYMM_0.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_YYYYMM_0.Focus();
                return;
            }
            if (iString.ISNull(WAGE_TYPE_0.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WAGE_TYPE_NAME_0.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0504_SET = new HRMF0504_SET(isAppInterfaceAdv1.AppInterface, "CLOSED_CANCEL"
                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                , PAY_YYYYMM_0.EditValue
                                                , WAGE_TYPE_0.EditValue, WAGE_TYPE_NAME_0.EditValue
                                                , DEPT_ID_0.EditValue, DEPT_CODE_0.EditValue, DEPT_NAME_0.EditValue
                                                , FLOOR_ID_0.EditValue, FLOOR_NAME_0.EditValue
                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue
                                                , CORP_TYPE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0504_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0504_SET.ShowDialog();
            vHRMF0504_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaMONTH_PAYMENT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_PAYMENT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ALLOWANCE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ALLOWANCE_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10108"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_ALLOWANCE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["ALLOWANCE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEDUCTION_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10109"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEDUCTION_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10108"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_DEDUCTION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (iString.ISNull(e.Row["MONTH_PAYMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10106"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaMONTH_PAYMENT_UpdateCompleted(object pSender)
        {
            Search_DB();
        }

        //// Pay Master 항목.
        //private void idaGRADE_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        //{
        //    if (e.Row["CORP_ID"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iString.ISNull(e.Row["START_YYYYMM"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Start Year Month(시작년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (e.Row["PERSON_ID"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person(사원)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (e.Row["PAY_TYPE"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Type(급여제)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //private void idaGRADE_HEADER_PreDelete(ISPreDeleteEventArgs e)
        //{
        //    if (e.Row.RowState != DataRowState.Added)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //// Allowance 항목.
        //private void idaPAY_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        //{
        //    if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //private void idaPAY_ALLOWANCE_PreDelete(ISPreDeleteEventArgs e)
        //{
        //    if (e.Row.RowState != DataRowState.Added)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //// Deduction 항목.
        //private void idaPAY_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        //{
        //    if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        //private void idaPAY_DEDUCTION_PreDelete(ISPreDeleteEventArgs e)
        //{
        //    if (e.Row.RowState != DataRowState.Added)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
        //        e.Cancel = true;
        //        return;
        //    }
        //}      
        #endregion

        #region ----- LookUp Event -----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("POST", "Y");
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }

        private void ilaWAGE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaALLOWANCE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("ALLOWANCE", "Y");
        }

        private void ilaDEDUCTION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("DEDUCTION", "Y"); 
        }

        #endregion

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }
    }
}