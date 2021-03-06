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

namespace HRMF0604
{
    public partial class HRMF0604_PAYMENT : Office2007Form
    {
       
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0604_PAYMENT(Form pMainForm, ISAppInterface pAppInterface
                                , object pYYYYMM, object pCORP_NAME, object pCORP_ID
                                , object pDEPT_NAME, object pDEPT_ID
                                , object pFLOOR_NAME, object pFLOOR_ID 
                                , object pNAME, object pPERSON_NUM, object pPERSON_ID)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_CORP_ID.EditValue = pCORP_ID;
            W_CORP_NAME.EditValue = pCORP_NAME;
            W_RESERVE_YYYYMM.EditValue = pYYYYMM;
            W_DEPT_ID.EditValue = pDEPT_ID;
            W_DEPT_NAME.EditValue = pDEPT_NAME;
            W_FLOOR_ID.EditValue = pFLOOR_ID;
            W_FLOOR_DESC.EditValue = pFLOOR_NAME; 
            W_PERSON_ID.EditValue = pPERSON_ID;
            W_PERSON_NUM.EditValue = pPERSON_NUM;
            W_NAME.EditValue = pNAME;
             
            V_PAYMENT_DATE.EditValue = iDate.ISMonth_Last(iDate.ISGetDate(string.Format("{0}-01", pYYYYMM)));
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {//업체
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (W_RESERVE_YYYYMM.EditValue == null)
            {//적용 년도  FCM_10036
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            IDA_RETIRE_RESERVE_PAY.Fill();
            IGR_RETIRE_RESERVE_PAY.Focus();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

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
            }
        }

        #endregion;

        #region ----- Form Event -----


        private void HRMF0604_PAYMENT_Shown(object sender, EventArgs e)
        {
            W_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
            W_PAYMENT_YN.EditValue = "N";
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void BTN_SEND_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 연말정산내역 급여전송 내용                        
            if (iString.ISNull(V_PAYMENT_DATE.EditValue) == String.Empty)
            {// 적용 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10445"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_PAYMENT_DATE.Focus();
                return;
            } 

            if (IGR_RETIRE_RESERVE_PAY.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            int vIDX_SELECT_YN = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("SELECT_YN");
            int vIDX_RESERVE_YYYYMM = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("RESERVE_YYYYMM");
            int vIDX_PERSON_ID = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("PERSON_ID");
            int vIDX_NAME = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("NAME"); 
            int vIDX_IRP_TYPE = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("IRP_TYPE");
             
            for (int vRow = 0; vRow < IGR_RETIRE_RESERVE_PAY.RowCount; vRow++)
            {
                if (IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_RETIRE_RESERVE_PAY.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);

                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_SELECT_YN", IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_SELECT_YN));
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_PERSON_ID", IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_RESERVE_YYYYMM", IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_RESERVE_YYYYMM));
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_IRP_TYPE", IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_IRP_TYPE));
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_PAYMENT_FLAG", W_PAYMENT_YN.EditValue);
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_PAYMENT_DATE", V_PAYMENT_DATE.EditValue);
                    IDC_SET_RETIRE_RESERVE_PAY.SetCommandParamValue("P_NAME", IGR_RETIRE_RESERVE_PAY.GetCellValue(vRow, vIDX_NAME)); 
                    IDC_SET_RETIRE_RESERVE_PAY.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_SET_RETIRE_RESERVE_PAY.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_SET_RETIRE_RESERVE_PAY.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_RETIRE_RESERVE_PAY.ExcuteError || mSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    IGR_RETIRE_RESERVE_PAY.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }

            IGR_RETIRE_RESERVE_PAY.LastConfirmChanges();
            IDA_RETIRE_RESERVE_PAY.OraSelectData.AcceptChanges();
            IDA_RETIRE_RESERVE_PAY.Refillable = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
        }

        private void BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        private void W_RB_NO_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;
            W_PAYMENT_YN.EditValue = vRadio.RadioCheckedString;
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_RETIRE_RESERVE_PAY.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_YN = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("SELECT_YN");
            for (int vRow = 0; vRow < IGR_RETIRE_RESERVE_PAY.RowCount; vRow++)
            {
                IGR_RETIRE_RESERVE_PAY.SetCellValue(vRow, vIDX_SELECT_YN, CB_SELECT_YN.CheckBoxValue);
            }

            IGR_RETIRE_RESERVE_PAY.LastConfirmChanges();
            IDA_RETIRE_RESERVE_PAY.OraSelectData.AcceptChanges();
            IDA_RETIRE_RESERVE_PAY.Refillable = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void IGR_YEAR_ADJUSTMENT__PAYMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_RETIRE_RESERVE_PAY.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_RETIRE_RESERVE_PAY.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_SELECT_YN)
            {
                IGR_RETIRE_RESERVE_PAY.LastConfirmChanges();
                IDA_RETIRE_RESERVE_PAY.OraSelectData.AcceptChanges();
                IDA_RETIRE_RESERVE_PAY.Refillable = true;
            }
        }

        #endregion              

        #region ----- Lookup Event -----
        
        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today));
        }

        private void ilaYEAR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYEAR.SetLookupParamValue("W_START_YEAR", "2001");
            ildYEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today));
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}