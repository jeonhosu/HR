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

namespace HRMF0636
{
    public partial class HRMF0636 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0636()
        {
            InitializeComponent();
        }

        public HRMF0636(Form pMainForm, ISAppInterface pAppInterface)
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
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();
                W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            if (itbBILL.SelectedTab.TabIndex == 1)
            {
                IDA_RETIRE_RESERVE.Fill();
                IGR_RETIRE_RESERVE.Focus();
            }
            else if (itbBILL.SelectedTab.TabIndex == 2)
            {
                IDA_RETIRE_RESERVE_SLIP.Fill();
                IGR_RETIRE_RESERVE_SLIP.Focus();
            }
            else if (itbBILL.SelectedTab.TabIndex == 3)
            {
                IDA_RETIRE_RESERVE_SLIP_CC.Fill();
                Summary_Amount();
                IGR_RETIRE_RESERVE_SLIP_CC.Focus();
            }
        }

        private void Summary_Amount()
        {
            int vIDX_DR_Amount = IGR_RETIRE_RESERVE_SLIP_CC.GetColumnToIndex("DR_AMOUNT");
            int vIDX_CR_Amount = IGR_RETIRE_RESERVE_SLIP_CC.GetColumnToIndex("CR_AMOUNT");

            decimal vDR_Amount = 0;
            decimal vCR_Amount = 0;
            decimal vGap_Amount = 0;

            for (int vRow = 0; vRow < IGR_RETIRE_RESERVE_SLIP_CC.RowCount; vRow++)
            {
                vDR_Amount = vDR_Amount + iConv.ISDecimaltoZero(IGR_RETIRE_RESERVE_SLIP_CC.GetCellValue(vRow, vIDX_DR_Amount));
                vCR_Amount = vCR_Amount + iConv.ISDecimaltoZero(IGR_RETIRE_RESERVE_SLIP_CC.GetCellValue(vRow, vIDX_CR_Amount));
            }
            vGap_Amount = Math.Abs(vDR_Amount - vCR_Amount) * -1;
            T_DR_SUM.EditValue = vDR_Amount;
            T_CR_SUM.EditValue = vCR_Amount;
            T_GAP_SUM.EditValue = vGap_Amount;
        }

        private void Set_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_SET_RETIRE_RESERVE_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_SET_RETIRE_RESERVE_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_SET_RETIRE_RESERVE_SLIP.GetCommandParamValue("O_MESSAGE"));
            
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_RETIRE_RESERVE_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }     
       
            //requery
            Search_DB();
        }

        private void Cancel_Slip_Interface()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vStatus = "F";
            string vMessage = string.Empty;

            IDC_CANCEL_RETIRE_RESERVE_SLIP.ExecuteNonQuery();
            vStatus = iConv.ISNull(IDC_CANCEL_RETIRE_RESERVE_SLIP.GetCommandParamValue("O_STATUS"));
            vMessage = iConv.ISNull(IDC_CANCEL_RETIRE_RESERVE_SLIP.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CANCEL_RETIRE_RESERVE_SLIP.ExcuteError || vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //requery
            Search_DB();
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
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_RETIRE_RESERVE.IsFocused)
                    {
                        IDA_RETIRE_RESERVE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0636_Shown(object sender, EventArgs e)
        {
            W_RESERVE_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            DefaultCorporation();              //Default Corp.
        }

        private void BTN_SET_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Slip_Interface();
        }

        private void BTN_CANCEL_SLIP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Cancel_Slip_Interface();
        }

        #endregion

        #region ----- Lookup Event ------

        #endregion

        #region ----- Adapter Event -----


        #endregion


    }
}