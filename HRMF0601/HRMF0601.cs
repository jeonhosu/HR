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

namespace HRMF0601
{
    public partial class HRMF0601 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----
        public HRMF0601(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Private Methods ----
        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorp()
        {
            //// Lookup SETTING
            //ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            //ildCORP.SetLookupParamValue("W_USABLE_CHECK_YN", "N");

            //// LOOKUP DEFAULT VALUE SETTING - CORP
            //idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            //idcDEFAULT_CORP.ExecuteNonQuery();

            //CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            //CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Init_Insert(int pSelectTabIndex)
        {
            if (pSelectTabIndex == 1)
            {
                STD_YYYY.EditValue = STD_YYYY_0.EditValue;
            }
            else if (pSelectTabIndex == 2)
            {
                igrCONTINUOUS_DEDUCTION.SetCellValue("STD_YYYY", STD_YYYY_0.EditValue);
            }
            else if (pSelectTabIndex == 3)
            {
                IGR_CHG_DEDUCTION_RATE.SetCellValue("ADJUST_YYYY", STD_YYYY_0.EditValue);
            }
        }

        private void SEARCH_DB()
        {
            if (STD_YYYY_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYY_0.Focus();
                return;
            }
            idaRETIRE_STANDARD.Fill();
            idaCONTINUOUS_DEDUCTION.Fill();
            IDA_CHG_DEDUCTION_RATE.Fill();

            if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 1)
            {
                STD_CALCULATE_MONTH.Focus();
            }
            else if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 2)
            {
                igrCONTINUOUS_DEDUCTION.Focus();
            }
            else if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 3)
            {
                IGR_CHG_DEDUCTION_RATE.Focus();
            }
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----
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
                    if (idaRETIRE_STANDARD.IsFocused)
                    {
                        idaRETIRE_STANDARD.AddOver();
                    }
                    else if(idaCONTINUOUS_DEDUCTION.IsFocused)
                    {
                        idaCONTINUOUS_DEDUCTION.AddOver();
                    }
                    else if (IDA_CHG_DEDUCTION_RATE.IsFocused)
                    {
                        IDA_CHG_DEDUCTION_RATE.AddOver();
                    }
                    Init_Insert(itbRETIRE_STANDARD.SelectedTab.TabIndex);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaRETIRE_STANDARD.IsFocused)
                    {
                        idaRETIRE_STANDARD.AddUnder();
                    }
                    else if (idaCONTINUOUS_DEDUCTION.IsFocused)
                    {
                        idaCONTINUOUS_DEDUCTION.AddUnder();
                    }
                    else if (IDA_CHG_DEDUCTION_RATE.IsFocused)
                    {
                        IDA_CHG_DEDUCTION_RATE.AddUnder();
                    }
                    Init_Insert(itbRETIRE_STANDARD.SelectedTab.TabIndex);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaRETIRE_STANDARD.IsFocused)
                    {
                        idaRETIRE_STANDARD.Update();
                    }
                    else if (idaCONTINUOUS_DEDUCTION.IsFocused)
                    {
                        idaCONTINUOUS_DEDUCTION.Update();
                    }
                    else if (IDA_CHG_DEDUCTION_RATE.IsFocused)
                    {
                        IDA_CHG_DEDUCTION_RATE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaRETIRE_STANDARD.IsFocused)
                    {
                        idaRETIRE_STANDARD.Cancel();
                    }
                    else if (idaCONTINUOUS_DEDUCTION.IsFocused)
                    {
                        idaCONTINUOUS_DEDUCTION.Cancel();
                    }
                    else if (IDA_CHG_DEDUCTION_RATE.IsFocused)
                    {
                        IDA_CHG_DEDUCTION_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaRETIRE_STANDARD.IsFocused)
                    {
                        idaRETIRE_STANDARD.Delete();
                    }
                    else if (idaCONTINUOUS_DEDUCTION.IsFocused)
                    {
                        idaCONTINUOUS_DEDUCTION.Delete();
                    }
                    else if (IDA_CHG_DEDUCTION_RATE.IsFocused)
                    {
                        IDA_CHG_DEDUCTION_RATE.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- From Event -----

        private void HRMF0601_Load(object sender, EventArgs e)
        {
            idaRETIRE_STANDARD.FillSchema();
            idaCONTINUOUS_DEDUCTION.FillSchema(); ;

            ildYEAR.SetLookupParamValue("W_START_YEAR", "2001");
            ildYEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today, 2));
            STD_YYYY_0.EditValue = iDate.ISYear(DateTime.Today);

            //DefaultCorp();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]          
        }

        private void ibtCOPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string mPre_YYYY;
            string mReturn_Value;
            DialogResult mDialogResult;

            if (iString.ISNull(STD_YYYY_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYY_0.Focus();
                return;
            }

            // 전년도 자료 존재 체크
            mPre_YYYY = Convert.ToString(Convert.ToInt32(STD_YYYY_0.EditValue) - Convert.ToInt32(1));
            if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 1)
            {// 기초정보.
                idcRETIRE_STANDARD_CHECK_YN.SetCommandParamValue("W_STD_YYYY", mPre_YYYY);
                idcRETIRE_STANDARD_CHECK_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(idcRETIRE_STANDARD_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "N".ToString())
                {// 기존 자료 존재.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10083"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    STD_YYYY_0.Focus();
                    return;
                }

                // 당년도 자료 존재 체크
                idcRETIRE_STANDARD_CHECK_YN.SetCommandParamValue("W_STD_YYYY", STD_YYYY_0.EditValue);
                idcRETIRE_STANDARD_CHECK_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(idcRETIRE_STANDARD_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "Y".ToString())
                {// 기존 자료 존재.
                    mDialogResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (mDialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                // Copy 시작.
                idcRETIRE_STANDARD_COPY.ExecuteNonQuery();
                string mSTATUS = idcRETIRE_STANDARD_COPY.GetCommandParamValue("O_STATUS").ToString();
                mReturn_Value = Convert.ToString(idcRETIRE_STANDARD_COPY.GetCommandParamValue("O_MESSAGE"));
                if (idcRETIRE_STANDARD_COPY.ExcuteError || mSTATUS == "F")
                {
                    if (mReturn_Value != string.Empty)
                    {
                        MessageBoxAdv.Show(mReturn_Value, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                MessageBoxAdv.Show(mReturn_Value, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 2)
            {// 근속 공제.
                idcCHECK_CONTINUOUS_DEDUCTION_YN.SetCommandParamValue("W_STD_YYYY", mPre_YYYY);
                idcCHECK_CONTINUOUS_DEDUCTION_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(idcCHECK_CONTINUOUS_DEDUCTION_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "N".ToString())
                {// 기존 자료 존재.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10083"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    STD_YYYY_0.Focus();
                    return;
                }

                // 당년도 자료 존재 체크
                idcCHECK_CONTINUOUS_DEDUCTION_YN.SetCommandParamValue("W_STD_YYYY", STD_YYYY_0.EditValue);
                idcCHECK_CONTINUOUS_DEDUCTION_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(idcCHECK_CONTINUOUS_DEDUCTION_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "Y".ToString())
                {// 기존 자료 존재.
                    mDialogResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (mDialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                // Copy 시작.
                idcCOPY_CONTINUOUS_DEDUCTION.ExecuteNonQuery();
                string mSTATUS = idcCOPY_CONTINUOUS_DEDUCTION.GetCommandParamValue("O_STATUS").ToString();
                mReturn_Value = Convert.ToString(idcCOPY_CONTINUOUS_DEDUCTION.GetCommandParamValue("O_MESSAGE"));
                if (idcCOPY_CONTINUOUS_DEDUCTION.ExcuteError || mSTATUS == "F")
                {
                    if (mReturn_Value != string.Empty)
                    {
                        MessageBoxAdv.Show(mReturn_Value, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                MessageBoxAdv.Show(mReturn_Value, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (itbRETIRE_STANDARD.SelectedTab.TabIndex == 3)
            {// 환산공제율관리.
                IDC_CHECK_CHG_DEDUCTION_RATE_YN.SetCommandParamValue("W_ADJUST_YYYY", mPre_YYYY);
                IDC_CHECK_CHG_DEDUCTION_RATE_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(IDC_CHECK_CHG_DEDUCTION_RATE_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "N".ToString())
                {// 기존 자료 존재.
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10083"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    STD_YYYY_0.Focus();
                    return;
                }

                // 당년도 자료 존재 체크
                IDC_CHECK_CHG_DEDUCTION_RATE_YN.SetCommandParamValue("W_ADJUST_YYYY", STD_YYYY_0.EditValue);
                IDC_CHECK_CHG_DEDUCTION_RATE_YN.ExecuteNonQuery();
                mReturn_Value = Convert.ToString(IDC_CHECK_CHG_DEDUCTION_RATE_YN.GetCommandParamValue("O_CHECK_YN"));
                if (mReturn_Value == "Y".ToString())
                {// 기존 자료 존재.
                    mDialogResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (mDialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                // Copy 시작.
                IDC_COPY_CHG_DEDUCTION_RATE.ExecuteNonQuery();
                string mSTATUS = IDC_COPY_CHG_DEDUCTION_RATE.GetCommandParamValue("O_STATUS").ToString();
                mReturn_Value = Convert.ToString(IDC_COPY_CHG_DEDUCTION_RATE.GetCommandParamValue("O_MESSAGE"));
                if (IDC_COPY_CHG_DEDUCTION_RATE.ExcuteError || mSTATUS == "F")
                {
                    if (mReturn_Value != string.Empty)
                    {
                        MessageBoxAdv.Show(mReturn_Value, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                MessageBoxAdv.Show(mReturn_Value, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion
        
        #region ----- Adapter Event -----
        private void idaRETIRE_STANDARD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["STD_YYYY"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaCONTINUOUS_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["STD_YYYY"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ----- Lookup Event -----

        private void ILA_RETIRE_AVG_AMT_CAL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "RETIRE_AVG_AMT_CAL");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y"); 
        }

        #endregion

    }
}