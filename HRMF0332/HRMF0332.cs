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

namespace HRMF0332
{
    public partial class HRMF0332 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0332()
        {
            InitializeComponent();
        }

        public HRMF0332(Form pMainForm, ISAppInterface pAppInterface)
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
            W_WORK_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_WORK_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (W_WORK_CORP_ID.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_CORP_NAME.Focus();
                return;
            }

            if (W_PERIOD_NAME.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_PERIOD_NAME.Focus();
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_SUMMARY.TabIndex)
            {
                IDA_OT_REQ_LIST.Fill();
            }
        }

        #endregion;


        #region ---- Set Grid Column(View, Hiden) ----- 

        private void Set_Grid_Col()
        {
            object vPeriod_Name = W_PERIOD_NAME.EditValue;

            //해당월 마지막 일자 이후는 모두 hiden 적용 
            int vLast_Day = iDate.ISMonth_Last(vPeriod_Name).Day;

            int vIDX_DAY_PERSON_29 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_PERSON_29");
            int vIDX_DAY_OT_29 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_OT_29");
            int vIDX_DAY_GAP_29 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_GAP_29");

            int vIDX_DAY_PERSON_30 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_PERSON_30");
            int vIDX_DAY_OT_30 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_OT_30");
            int vIDX_DAY_GAP_30 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_GAP_30");

            int vIDX_DAY_PERSON_31 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_PERSON_31");
            int vIDX_DAY_OT_31 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_OT_31");
            int vIDX_DAY_GAP_31 = IGR_OT_REQ_LIST.GetColumnToIndex("DAY_GAP_31");

            //Grid View/Hiden
            if (vLast_Day == 28)
            {
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_29].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_29].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_29].Visible = 0;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_30].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_30].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_30].Visible = 0;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_31].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_31].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_31].Visible = 0;
            }
            else if (vLast_Day == 29)
            {
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_29].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_29].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_29].Visible = 1;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_30].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_30].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_30].Visible = 0;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_31].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_31].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_31].Visible = 0;
            }
            else if (vLast_Day == 30)
            {
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_29].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_29].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_29].Visible = 1;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_30].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_30].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_30].Visible = 1;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_31].Visible = 0;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_31].Visible = 0;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_31].Visible = 0;
            }
            else 
            {
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_29].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_29].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_29].Visible = 1;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_30].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_30].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_30].Visible = 1;

                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_PERSON_31].Visible = 1;
                //IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_OT_31].Visible = 1;
                IGR_OT_REQ_LIST.GridAdvExColElement[vIDX_DAY_GAP_31].Visible = 1;
            }
            IGR_OT_REQ_LIST.ResetDraw = true;
        }

        #endregion

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
                    if (IDA_OT_REQ_LIST.IsFocused)
                    {
                        IDA_OT_REQ_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OT_REQ_LIST.IsFocused)
                    {
                        IDA_OT_REQ_LIST.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0332_Load(object sender, EventArgs e)
        {
            // Year Month Setting
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
            W_PERIOD_NAME.EditValue = iDate.ISYearMonth(DateTime.Today);
            // CORP SETTING
            DefaultCorporation();
            
            //그리드 제어 
            Set_Grid_Col();

            IDA_OT_REQ_LIST.FillSchema();

        }

        #endregion

        #region ----- Lookup Event ----- 
        
        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
        }

        #endregion

        private void ILA_W_YYYYMM_SelectedRowData(object pSender)
        {
            Set_Grid_Col();
        }

    }
}