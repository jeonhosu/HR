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

namespace HRMF0722
{
    public partial class HRMF0722 : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public HRMF0722()
        {
            InitializeComponent();
        }

        public HRMF0722(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YEAR_YYYY.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show("정산년도는 필수입니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYY.Focus();
                return;
            }
            if (TB_MAIN.SelectedTab.TabIndex == TP_DIST.TabIndex)
            {
                IDA_YEAR_ADJUST_DIST_LIST.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_DIST_SP.TabIndex)
            {
                IDA_YEAR_ADJUST_DIST_DTL_I.Fill();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_PAYMENT_SP.TabIndex)
            {
                IDA_YEAR_ADJUST_DIST_DTL_II.Fill();
            }
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
                    if (MessageBoxAdv.Show("해당 자료를 삭제하시겠습니까?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }

                    try
                    {
                        IDC_DELETE_YEAR_ADJUST_DIST.ExecuteNonQuery();

                        Search_DB();
                    }
                    catch (Exception Ex)
                    {
                        MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event ------
        
        private void HRMF0722_Load(object sender, EventArgs e)
        {
            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            W_CORP_NAME.BringToFront();

            // Standard Date SETTING
            //DateTime dLastYearMonthDay = new DateTime(DateTime.Today.Year, 12, 31);
            //STD_YYYYMM.EditValue = dLastYearMonthDay;
            W_YEAR_YYYY.EditValue = iDate.ISYear(iDate.ISDate_Month_Add(DateTime.Today, -12));

            IDA_YEAR_ADJUST_DIST_LIST.FillSchema(); 
        }

        private void BTN_ADJUST_DIST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_YEAR_YYYY.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYY.Focus();
                return;
            }

            HRMF0722_DIST vHRMF0722_DIST = new HRMF0722_DIST(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                            , W_YEAR_YYYY.EditValue
                                                            , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                            , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                            , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue
                                                            , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0722_DIST.ShowDialog();
            vHRMF0722_DIST.Dispose();
        }

        private void btnPayTransfer_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_YEAR_YYYY.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYY.Focus();
                return;
            }

            HRMF0722_SALARY vHRMF0722_SALARY = new HRMF0722_SALARY(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                    , W_YEAR_YYYY.EditValue 
                                                                    , W_CORP_NAME.EditValue, W_CORP_ID.EditValue
                                                                    , W_DEPT_NAME.EditValue, W_DEPT_ID.EditValue
                                                                    , W_FLOOR_DESC.EditValue, W_FLOOR_ID.EditValue                                                                     
                                                                    , W_PERSON_NAME.EditValue, W_PERSON_NUM.EditValue, W_PERSON_ID.EditValue);
            vHRMF0722_SALARY.ShowDialog();
            vHRMF0722_SALARY.Dispose();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion




    }
}