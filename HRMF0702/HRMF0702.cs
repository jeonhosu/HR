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

namespace HRMF0702
{
    public partial class HRMF0702 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----

        #endregion;

        #region ----- Constructor -----

        public HRMF0702()
        {
            InitializeComponent();
        }

        public HRMF0702(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }

        private bool Check_Prework_Added()
        {
            Boolean Row_Added_Status = false;
            //헤더 체크 
            for (int r = 0; r < idaPREVIOUS_WORK_LIST.SelectRows.Count; r++)
            {
                if (idaPREVIOUS_WORK_LIST.SelectRows[r].RowState == DataRowState.Added ||
                    idaPREVIOUS_WORK_LIST.SelectRows[r].RowState == DataRowState.Modified)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10169"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }             
            return (Row_Added_Status);
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    // 조회 전, 조건 체크하는 부분
                    if (iString.ISNull(CORP_NAME_0.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        CORP_NAME_0.Focus();
                        return;
                    }
                    if (iString.ISNull(STD_YYYYMM.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        STD_YYYYMM.Focus();
                        return;
                    }

                    // Person Info.
                    idaPERSON_INFO.Fill();

                    idaPREVIOUS_WORK_LIST.Fill();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    // 새로운 데이터 등록 시, 기준 사용자 조회 체크
                    if (iString.ISNull(PERSON_NAME_1.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10058"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        W_PERSON_NAME.Focus();
                        return;
                    }

                    if (Check_Prework_Added() == true)
                    {
                        return;
                    }

                    idaPREVIOUS_WORK_LIST.AddOver();
                    IDC_SEQ_NUMBER.ExecuteNonQuery();
               
                   
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    // 새로운 데이터 등록 시, 기준 사용자 조회 체크
                    if (iString.ISNull(PERSON_NAME_1.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10058"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        W_PERSON_NAME.Focus();
                        return;
                    }

                    if (Check_Prework_Added() == true)
                    {
                        return;
                    }

                    idaPREVIOUS_WORK_LIST.AddUnder();
                    IDC_SEQ_NUMBER.ExecuteNonQuery();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaPERSON_INFO.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {

                    idaPREVIOUS_WORK_LIST.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                    idaPREVIOUS_WORK_LIST.Delete();
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0702_Load(object sender, EventArgs e)
        {
            //idaPERSON_INFO.FillSchema();

            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();

            if (DateTime.Today.Month <= 2)
            {
                STD_YYYYMM.EditValue = iDate.ISYearMonth(iDate.ISDate_Add(string.Format("{0}-01-01", DateTime.Today.Year), -1));
            }
            else
            {
                STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            }
            idaPERSON_INFO.FillSchema();

            W_PERSON_NAME.Focus();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 1)));
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("YEAR_EMPLOYE_TYPE", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        #endregion

        #region ----- Adapter Event ------

        private void idaSEQ_ONE_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iDate.ISGetDate(e.Row["JOIN_DATE"]) > iDate.ISGetDate(e.Row["RETR_DATE"]))
            {
                MessageBoxAdv.Show("퇴직일자보다 입사일자가 클 수 없습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["RETR_DATE"]).Year != iString.ISNumtoZero(YEAR_0.EditValue))
            {
                MessageBoxAdv.Show("퇴직일자는 당해년도 이내여야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaSEQ_TWO_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iDate.ISGetDate(e.Row["JOIN_DATE"]) > iDate.ISGetDate(e.Row["RETR_DATE"]))
            {
                MessageBoxAdv.Show("퇴직일자보다 입사일자가 클 수 없습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iDate.ISGetDate(e.Row["RETR_DATE"]).Year != iString.ISNumtoZero(YEAR_0.EditValue))
            {
                MessageBoxAdv.Show("퇴직일자는 당해년도 이내여야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}