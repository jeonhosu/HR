using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace HRMF0379
{
    public partial class HRMF0379 : Office2007Form
    {
        #region ----- Variables -----

        DateTime vSYS_DATE = DateTime.Today;
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0379()
        {
            InitializeComponent();
        }

        public HRMF0379(Form pMainForm, ISAppInterface pAppInterface)
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
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            
            WORK_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            WORK_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (WORK_CORP_ID_0.EditValue == null)
            {// 업체. - 업체정보는 필수입니다. 선택하세요.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자 - 시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }

            IDC_GET_LOCAL_DATE_TIME_P.ExecuteNonQuery();
            vSYS_DATE = iDate.ISGetDate(IDC_GET_LOCAL_DATE_TIME_P.GetCommandParamValue("X_LOCAL_DATE"));

            IDA_DAY_IF_SUMMARY_I.SetSelectParamValue("W_SYS_DATE", vSYS_DATE);
            IDA_DAY_IF_SUMMARY_I.Fill();

            IDA_DAY_IF_LIST_C.SetSelectParamValue("W_SYS_DATE", vSYS_DATE);
            IDA_DAY_IF_LIST_C.Fill();

            IDA_DAY_IF_LIST_D.SetSelectParamValue("W_SYS_DATE", vSYS_DATE);
            IDA_DAY_IF_LIST_D.Fill();

            IDC_GET_DAY_IF_JOIN_RETIRE_LIST.SetCommandParamValue("W_SYS_DATE", vSYS_DATE);
            IDC_GET_DAY_IF_JOIN_RETIRE_LIST.ExecuteNonQuery();
            O_JOIN_10_LIST.EditValue = IDC_GET_DAY_IF_JOIN_RETIRE_LIST.GetCommandParamValue("O_JOIN_10_LIST");
            O_JOIN_ETC_LIST.EditValue = IDC_GET_DAY_IF_JOIN_RETIRE_LIST.GetCommandParamValue("O_JOIN_ETC_LIST");
            O_RETIRE_NEW_LIST.EditValue = IDC_GET_DAY_IF_JOIN_RETIRE_LIST.GetCommandParamValue("O_RETIRE_NEW_LIST");
            O_RETIRE_10_LIST.EditValue = IDC_GET_DAY_IF_JOIN_RETIRE_LIST.GetCommandParamValue("O_RETIRE_10_LIST");
            O_RETIRE_ETC_LIST.EditValue = IDC_GET_DAY_IF_JOIN_RETIRE_LIST.GetCommandParamValue("O_RETIRE_ETC_LIST"); 
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
                    if (IDA_DAY_IF_SUMMARY_I.IsFocused)
                    {
                        IDA_DAY_IF_SUMMARY_I.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DAY_IF_SUMMARY_I.IsFocused)
                    {
                        IDA_DAY_IF_SUMMARY_I.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0379_Load(object sender, EventArgs e)
        {

        }

        private void HRMF0379_Shown(object sender, EventArgs e)
        {
            WORK_DATE_0.EditValue = DateTime.Today;
            W_PRE_HOLY_3_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;

            DefaultCorporation();

            WORK_DATE_0.Focus();

            RB_WORK_CENTER.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_SORT_FLAG.EditValue = "WC";

        }

        private void RB_WORK_CENTER_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;
            if (vRadio.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                W_SORT_FLAG.EditValue = vRadio.RadioCheckedString;
            }
        }

        private void TB_MAIN_Click(object sender, EventArgs e)
        {
            //Search_DB();
        }

        private void IGR_DAY_IF_SUMMARY_I_CellDoubleClick(object pSender)
        {
            int vIDX_COL = IGR_DAY_IF_SUMMARY_I.ColIndex;
            int vIDX_ROW = IGR_DAY_IF_SUMMARY_I.RowIndex;
            if (vIDX_COL < 0 || vIDX_ROW < 0)
            {
                return;
            }
             
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            object vDEPT_ID = IGR_DAY_IF_SUMMARY_I.GetCellValue(vIDX_ROW, IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DEPT_ID"));
            object vSUB_TYPE = IGR_DAY_IF_SUMMARY_I.GetCellValue(vIDX_ROW, IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("SUB_TYPE"));
            object vDUTY_TYPE = string.Empty;
            if(vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("LATE_COUNT"))
            {
                //지각//
                vDUTY_TYPE = "LATE";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("ABSENT_COUNT"))
            {
                //결근//
                vDUTY_TYPE = "ABSENT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_1_COUNT"))
            {
                //연차//
                vDUTY_TYPE = "DUTY_1_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_2_COUNT"))
            {
                //연차//
                vDUTY_TYPE = "DUTY_2_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_3_COUNT"))
            {
                //경조//
                vDUTY_TYPE = "DUTY_3_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_4_COUNT"))
            {
                //교육//
                vDUTY_TYPE = "DUTY_4_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_5_COUNT"))
            {
                //무급//
                vDUTY_TYPE = "DUTY_5_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_6_COUNT"))
            {
                //출산/육아//
                vDUTY_TYPE = "DUTY_6_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_18_COUNT"))
            {
                //기타//
                vDUTY_TYPE = "DUTY_18_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("DUTY_19_COUNT"))
            {
                //휴무//
                vDUTY_TYPE = "DUTY_19_COUNT";
            }
            else if (vIDX_COL == IGR_DAY_IF_SUMMARY_I.GetColumnToIndex("RETIRE_COUNT"))
            {
                //퇴직//
                vDUTY_TYPE = "RETIRE_COUNT";
            } 
            IDA_DAY_IF_LIST.SetSelectParamValue("W_DEPT_ID", vDEPT_ID);
            IDA_DAY_IF_LIST.SetSelectParamValue("W_SUB_TYPE", vSUB_TYPE);
            IDA_DAY_IF_LIST.SetSelectParamValue("W_DUTY_TYPE", vDUTY_TYPE);
            IDA_DAY_IF_LIST.SetSelectParamValue("W_SYS_DATE", vSYS_DATE);
            IDA_DAY_IF_LIST.Fill(); 
        }

        #endregion

        #region ----- Lookup event -----

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}