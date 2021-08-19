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

namespace HRMF0356
{
    public partial class HRMF0356 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0356()
        {
            InitializeComponent();
        }

        public HRMF0356(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_DESC_2.EditValue = W_CORP_NAME_0.EditValue;
            W_CORP_ID_2.EditValue = W_CORP_ID_0.EditValue;

            IDC_DEFAULT_CLOSING_TYPE.ExecuteNonQuery();
            W_DUTY_TYPE_DESC_2.EditValue = IDC_DEFAULT_CLOSING_TYPE.GetCommandParamValue("O_CODE_NAME");
            W_DUTY_TYPE_2.EditValue = IDC_DEFAULT_CLOSING_TYPE.GetCommandParamValue("O_CODE");
        }

        private void SEARCH_DB()
        {
            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                IDA_DAY_LEAVE_TOTAL.Fill();
            }
            else if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                IDA_MONTH_TOTAL_PERIOD.Fill();
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
                    SEARCH_DB();
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (IDA_DAY_LEAVE_TOTAL.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (IDA_DAY_LEAVE_TOTAL.IsFocused)
                    {
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        
        private void HRMF0356_Load(object sender, EventArgs e)
        {
            W_START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_END_DATE_0.EditValue = DateTime.Today;

            W_DUTY_YYYYMM_FR_2.EditValue = iDate.ISYearMonth(DateTime.Today);
            W_DUTY_YYYYMM_TO_2.EditValue = W_DUTY_YYYYMM_FR_2.EditValue;

            DefaultCorporation();

            IDA_DAY_LEAVE_TOTAL.FillSchema();            
        }

        private void HRMF0356_Shown(object sender, EventArgs e)
        {
            
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //ildPERSON.SetLookupParamValue("W_END_DATE", END_DATE.EditValue);
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_CORP_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_DUTY_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_TYPE_2.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_FLOOR_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_PERSON_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_YYYYMM_FR_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", string.Empty);
        }

        private void ILA_YYYYMM_TO_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", W_DUTY_YYYYMM_FR_2.EditValue);
        }

        #endregion

        

       
    }
}