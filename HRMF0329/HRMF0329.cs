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

namespace HRMF0329
{
    public partial class HRMF0329 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0329()
        {
            InitializeComponent();
        }

        public HRMF0329(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");


            // LEAVE CLOSE TYPE SETTING
            ILD_CLOSED_FLAG_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ILD_CLOSED_FLAG_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            CLOSED_FLAG_NAME_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            CLOSED_FLAG_0.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WEEK_START_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WEEK_START_DATE_0.Focus();
                return;
            }
            if (WEEK_END_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WEEK_END_DATE_0.Focus();
                return;
            }

            IDA_PERSON_WEEK_40H.Fill();
            IGR_PERSON_WEEK_40H.Focus();
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_PERSON_WEEK_40H.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_PERSON_WEEK_40H.Cancel();
                    IDA_DAY_LEAVE_WEEK.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0329_Load(object sender, EventArgs e)
        {
            DefaultValues();

            DUTY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            IDA_PERSON_WEEK_40H.FillSchema();
        }

        private void HRMF0329_Shown(object sender, EventArgs e)
        {
            
        }

        #endregion

        #region ----- Lookup event ------

        private void ilaYYYYMM_0_SelectedRowData(object pSender)
        {
            WEEK_CODE_0.EditValue = null;
            WEEK_START_DATE_0.EditValue = null;
            WEEK_END_DATE_0.EditValue = null;
        }

        private void ILA_PERSON_0_SelectedRowData(object pSender)
        {
            Search_DB();
        }

        private void ILA_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_HOLY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("HOLY_TYPE", "Y");
        }

        private void ILA_WORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("WORK_TYPE", "Y");
        }

        private void ILA_DUTY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DUTY", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ILA_JOB_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("JOB_CATEGORY", "Y");
        }

        private void ILA_DUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("DUTY", "Y");
        }

        private void ILA_HOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("HOLY_TYPE", "Y");
        }

        private void ILA_YYYYMM_WEEK_SelectedRowData(object pSender)
        {
            idcYYYYMM_WEEK.SetCommandParamValue("W_WEEK_CODE", WEEK_CODE_0.EditValue);
            idcYYYYMM_WEEK.ExecuteNonQuery();
        }
        #endregion

        #region ------ Adpater event ------

        private void IDA_DAY_LEAVE_WEEK_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["DAY_LEAVE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10471"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DUTY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10175"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["HOLY_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10470"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            //근태코드 && 근무구분 상호 검증 //
            //1.근태 (정상/결근) => 근무구분(주간/야간) 선택가능.
            if (iConv.ISNull(e.Row["DUTY_CODE"]) == "00" || iConv.ISNull(e.Row["DUTY_CODE"]) == "11")
            {
                if (iConv.ISNull(e.Row["HOLY_TYPE"]) == "2" || iConv.ISNull(e.Row["HOLY_TYPE"]) == "3")
                {
                }
                else
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10472", string.Format("&&DUTY_NAME:={0}&&HOLY_TYPE:={1}", "출근/결근", "주간/야간")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }
            }
            //2.근태(무급휴일) => 근무구분(무휴) 선택가능.
            if (iConv.ISNull(e.Row["DUTY_CODE"]) == "52" )
            {
                if (iConv.ISNull(e.Row["HOLY_TYPE"]) == "0")
                {
                }
                else
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10472", string.Format("&&DUTY_NAME:={0}&&HOLY_TYPE:={1}", "무급휴일", "무휴")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }
            }
            //3.근태(유급휴일) => 근무구분(유휴) 선택가능.
            if (iConv.ISNull(e.Row["DUTY_CODE"]) == "51")
            {
                if (iConv.ISNull(e.Row["HOLY_TYPE"]) == "1")
                {
                }
                else
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10472", string.Format("&&DUTY_NAME:={0}&&HOLY_TYPE:={1}", "유급휴일", "유휴")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }
            }
            //4.근태(휴일근무) => 근무구분(무휴/유휴) 선택가능.
            if (iConv.ISNull(e.Row["DUTY_CODE"]) == "53")
            {
                if (iConv.ISNull(e.Row["HOLY_TYPE"]) == "0" || iConv.ISNull(e.Row["HOLY_TYPE"]) == "1")
                {
                }
                else
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10472", string.Format("&&DUTY_NAME:={0}&&HOLY_TYPE:={1}", "휴일근무", "유휴/무휴")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }
            }
        }

        #endregion



    }
}