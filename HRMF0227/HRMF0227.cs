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

namespace HRMF0227
{
    public partial class HRMF0227 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0227()
        {
            InitializeComponent();
        }

        public HRMF0227(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
           
        }

        private void Search_DB()
        {
            if (TB_BASE.SelectedTab.TabIndex == 1)
            {
                IDA_HRM_CAR.Fill();
                IGR_HRM_CAR.Focus();
            }
            if (TB_BASE.SelectedTab.TabIndex == 2)
            {
                IDA_CAR_TYPE.Fill();
                IGR_CAR_TYPE.Focus();
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
                    if (IDA_HRM_CAR.IsFocused)
                    {
                        IDA_HRM_CAR.AddOver();

                        IGR_HRM_CAR.SetCellValue("ENABLED_FLAG", "Y");
                        IGR_HRM_CAR.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                    }
                    if (IDA_CAR_TYPE.IsFocused)
                    {
                        IDA_CAR_TYPE.AddOver();

                        IGR_CAR_TYPE.SetCellValue("ENABLED_FLAG", "Y");
                        IGR_CAR_TYPE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                    }
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_HRM_CAR.IsFocused)
                    {
                        IDA_HRM_CAR.AddUnder();

                        IGR_HRM_CAR.SetCellValue("ENABLED_FLAG", "Y");
                        IGR_HRM_CAR.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                    }
                    if (IDA_CAR_TYPE.IsFocused)
                    {
                        IDA_CAR_TYPE.AddUnder();

                        IGR_CAR_TYPE.SetCellValue("ENABLED_FLAG", "Y");
                        IGR_CAR_TYPE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_HRM_CAR.IsFocused)
                    {
                        IDA_HRM_CAR.Update();
                    }
                    if (IDA_CAR_TYPE.IsFocused)
                    {
                        IDA_CAR_TYPE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_HRM_CAR.IsFocused)
                    {
                        IDA_HRM_CAR.Cancel();
                    }
                    if (IDA_CAR_TYPE.IsFocused)
                    {
                        IDA_CAR_TYPE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_HRM_CAR.IsFocused)
                    {
                        IDA_HRM_CAR.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Events -----

        private void HRMF0227_Load(object sender, EventArgs e)
        {
            IDA_HRM_CAR.FillSchema();
            IDA_CAR_TYPE.FillSchema();
        }

        private void HRMF0227_Shown(object sender, EventArgs e)
        {
            DefaultValues();

            STD_DATE_0.EditValue = DateTime.Today;
            CB_ENABLE_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
             

        }

        private void ILA_REALATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", "RELATION");
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", "VALUE2 = 'Y'");
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

        }

        private void ILA_CAR_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "CAR_TYPE");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
        }

        private void ILA_OIL_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "OIL_TYPE");
        }

        private void ILA_GEARBOX_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "GEARBOX_TYPE");
        }
        #endregion;

        

        
        

    }
}