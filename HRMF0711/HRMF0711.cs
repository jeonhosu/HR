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

namespace HRMF0711
{
    public partial class HRMF0711 : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0711()
        {
            InitializeComponent();
        }

        public HRMF0711(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                idcDEFAULT_CORP.ExecuteNonQuery();
                W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                W_CORP_NAME.BringToFront();
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        // 업체정보, 기준일자, 해당 년도 값 체크 후 - Person Info 그리드에 데이터 출력
        private void SearchDB()
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("정산년도는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR.Focus();
                return;
            }
            idaPERSON.Fill();
            igrPERSON_INFO.Focus();
        }

        private void Init_Medical_Sum()
        {
            IDC_MEDICAL_INFO_SUM_P.ExecuteNonQuery();
            O_CREDIT_COUNT.EditValue = IDC_MEDICAL_INFO_SUM_P.GetCommandParamValue("O_CREDIT_COUNT");
            O_CREDIT_AMT.EditValue = IDC_MEDICAL_INFO_SUM_P.GetCommandParamValue("O_CREDIT_AMT");
            O_ETC_COUNT.EditValue = IDC_MEDICAL_INFO_SUM_P.GetCommandParamValue("O_ETC_COUNT");
            O_ETC_AMT.EditValue = IDC_MEDICAL_INFO_SUM_P.GetCommandParamValue("O_ETC_AMT");
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
                    idaMEDICAL_INFO.AddOver();
                    igrMEDIC_INFO.Focus();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    idaMEDICAL_INFO.AddUnder();
                    igrMEDIC_INFO.Focus();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaPERSON.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaMEDICAL_INFO.IsFocused)
                    {
                        idaMEDICAL_INFO.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaMEDICAL_INFO.IsFocused)
                    {
                        idaMEDICAL_INFO.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ------ form event -----

        private void HRMF0711_Load(object sender, EventArgs e)
        {
            idaPERSON.FillSchema();
            idaMEDICAL_INFO.FillSchema();
        }

        private void HRMF0711_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            // Standard Date SETTING
            DateTime dLastYearMonthDay = DateTime.Today;
            if (DateTime.Today.Month <= 2)
            {
                dLastYearMonthDay = new DateTime(DateTime.Today.AddYears(-1).Year, 12, 31);
            }
            else
            {
                dLastYearMonthDay = new DateTime(DateTime.Today.Year, 12, 31);
            }
            W_YEAR.EditValue = dLastYearMonthDay.Year;
        }

        private void BTN_INIT_YESONE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YEAR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("정산년도는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            IDC_INIT_MEDICAL_INFO.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_INIT_MEDICAL_INFO.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_INIT_MEDICAL_INFO.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            if (IDC_INIT_MEDICAL_INFO.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            SearchDB();
        }

        #endregion

        #region ------ Lookup Event ------

        private void ilaMEDIC_EVIDENCE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "MEDIC_EVIDENCE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaMEDIC_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "MEDIC_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            int vYEAR = iString.ISNumtoZero(W_YEAR.EditValue);
            DateTime vSTART_DATE = new DateTime(vYEAR, 1, 1);
            DateTime vEND_DATE = new DateTime(vYEAR, 12, 31);
            ildPERSON_0.SetLookupParamValue("W_START_DATE", vSTART_DATE);
            ildPERSON_0.SetLookupParamValue("W_END_DATE", vEND_DATE);
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- ADAPTER EVENT -----

        private void idaMEDICAL_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(W_YEAR.EditValue) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("정산년도가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("사원정보 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["RELATION_CODE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("관계정보 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["FAMILY_NAME"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("성명은 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("주민번호는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(e.Row["EVIDENCE_CODE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("의료비 증빙코드는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["CREDIT_AMT"], 0) + iString.ISDecimaltoZero(e.Row["ETC_AMT"], 0) == 0)
            {
                e.Cancel = true;
                MessageBoxAdv.Show("지급금액은 0보다 커야 합니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void idaMEDICAL_INFO_FilterCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            Init_Medical_Sum();
        }

        private void idaMEDICAL_INFO_UpdateCompleted(object pSender)
        {
            Init_Medical_Sum();
        }

        #endregion

    }
}