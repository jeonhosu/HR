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


namespace HRMF0521
{
    public partial class HRMF0521 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public HRMF0521()
        {
            InitializeComponent();
        }

        public HRMF0521(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
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

            idaSALARY_OUTSOURCING.Fill();
            igrSALARY_OUTSOURCING.Focus();
            //igrPERSON.Focus();
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
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
                    if (idaSALARY_OUTSOURCING.IsFocused)
                    {
                        idaSALARY_OUTSOURCING.AddOver();

                        igrSALARY_OUTSOURCING.SetCellValue("PAY_YYYYMM", PAY_YYYYMM_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("WAGE_TYPE", WAGE_TYPE_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("WAGE_TYPE_NAME", WAGE_TYPE_NAME_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("DEPT_ID", DEPT_ID_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("DEPT_NAME", DEPT_NAME_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("JOB_CATEGORY_ID", JOB_CATEGORY_ID.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("JOB_CATEGORY_NAME", JOB_CATEGORY_NAME.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaSALARY_OUTSOURCING.IsFocused)
                    {
                        idaSALARY_OUTSOURCING.AddUnder();

                        igrSALARY_OUTSOURCING.SetCellValue("PAY_YYYYMM", PAY_YYYYMM_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("WAGE_TYPE", WAGE_TYPE_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("WAGE_TYPE_NAME", WAGE_TYPE_NAME_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("DEPT_ID", DEPT_ID_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("DEPT_NAME", DEPT_NAME_0.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("JOB_CATEGORY_ID", JOB_CATEGORY_ID.EditValue);
                        igrSALARY_OUTSOURCING.SetCellValue("JOB_CATEGORY_NAME", JOB_CATEGORY_NAME.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaSALARY_OUTSOURCING.IsFocused)
                    {
                        idaSALARY_OUTSOURCING.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSALARY_OUTSOURCING.IsFocused)
                    {
                        idaSALARY_OUTSOURCING.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaSALARY_OUTSOURCING.IsFocused)
                    {
                        idaSALARY_OUTSOURCING.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0521_Load(object sender, EventArgs e)
        {
            idaSALARY_OUTSOURCING.FillSchema();
        }

        private void HRMF0521_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();              //Default Corp.

            PAY_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = iDate.ISMonth_Last(DateTime.Today);
        }

        #endregion;

        #region ----- LookUp Event -----

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaJOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWAGE_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_CONTROL_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDUTY_CONTROL.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildDUTY_CONTROL.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaJOB_CATEGORY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion;

        private void idaSALARY_OUTSOURCING_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10020"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["FLOOR_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10017"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["JOB_CATEGORY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10481"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["PERSON_COUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=인원"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["TOTAL_SUPPLY_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=지급액"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }


        }

        private void JOB_CATEGORY_NAME_Load(object sender, EventArgs e)
        {


        }

    }
}