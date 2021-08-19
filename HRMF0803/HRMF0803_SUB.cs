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

namespace HRMF0803
{
    public partial class HRMF0803_SUB : Office2007Form
    {
       
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0803_SUB(Form pMainForm, ISAppInterface pAppInterface
                                , object pYYYYMM, object pSTART_DATE, object pEND_DATE
                                , object pCORP_NAME, object pCORP_ID
                                , object pDEPT_NAME, object pDEPT_ID  
                                , object pNAME, object pPERSON_NUM, object pPERSON_ID
                                , string pSTATUS)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_CORP_ID.EditValue = pCORP_ID;
            W_CORP_NAME.EditValue = pCORP_NAME;
            W_DEPT_ID.EditValue = pDEPT_ID;
            W_DEPT_NAME.EditValue = pDEPT_NAME;  
            W_PERSON_ID.EditValue = pPERSON_ID;
            W_PERSON_NUM.EditValue = pPERSON_NUM;
            W_NAME.EditValue = pNAME;

            W_FOOD_YYYYMM.EditValue = pYYYYMM;
            W_START_DATE.EditValue = pSTART_DATE;
            W_END_DATE.EditValue = pEND_DATE;

            P_PAY_YYYYMM.EditValue = pYYYYMM;

            if (pSTATUS == "CANCEL")
            {
                W_RB_YES.CheckedState = ISUtil.Enum.CheckedState.Checked;
                W_TRANS_YN.EditValue = W_RB_YES.RadioCheckedString;
                BTN_SEND.Visible = false;
                BTN_CANCEL.Visible = true;
            }
            else
            {
                W_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
                W_TRANS_YN.EditValue = W_RB_NO.RadioCheckedString;
                BTN_SEND.Visible = true;
                BTN_CANCEL.Visible = false;
            }
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
            if (W_FOOD_YYYYMM.EditValue == null)
            {//적용 년도  FCM_10036
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_FOOD_YYYYMM.Focus();
                return;
            }

            CB_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked; 
            IDA_FOOD_DED_PAYMENT.Fill();
            IGR_FOOD_DED_PAYMENT.Focus();
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

        private void HRMF0803_SUB_Load(object sender, EventArgs e)
        {
           
        }

        private void HRMF0803_SUB_Shown(object sender, EventArgs e)
        {
            
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void BTN_SEND_ButtonClick(object pSender, EventArgs pEventArgs)
        {                   
            if (iString.ISNull(P_PAY_YYYYMM.EditValue) == String.Empty)
            {// 적용 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                P_PAY_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(P_WAGE_TYPE.EditValue) == string.Empty)
            {//급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                P_WAGE_TYPE_NAME.Focus();
                return;
            }

            if (IGR_FOOD_DED_PAYMENT.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            int vIDX_SELECT_YN = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("SELECT_YN");
            int vIDX_FOOD_DED_AMOUNT = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("FOOD_DED_AMOUNT");
            int vIDX_PERSON_ID = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("PERSON_ID"); 
             
            for (int vRow = 0; vRow < IGR_FOOD_DED_PAYMENT.RowCount; vRow++)
            {
                if (IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_FOOD_DED_PAYMENT.CurrentCellMoveTo(vRow, vIDX_SELECT_YN); 

                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_TRANS_STATUS", W_TRANS_YN.EditValue);
                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_PERSON_ID", IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_FOOD_DED_AMOUNT", IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_FOOD_DED_AMOUNT));
                    IDC_EXEC_FOOD_DED_PAYMENT.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_EXEC_FOOD_DED_PAYMENT.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_EXEC_FOOD_DED_PAYMENT.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_EXEC_FOOD_DED_PAYMENT.ExcuteError || mSTATUS == "F")
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
                    IGR_FOOD_DED_PAYMENT.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }

            IGR_FOOD_DED_PAYMENT.LastConfirmChanges();
            IDA_FOOD_DED_PAYMENT.OraSelectData.AcceptChanges();
            IDA_FOOD_DED_PAYMENT.Refillable = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(P_PAY_YYYYMM.EditValue) == String.Empty)
            {// 적용 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                P_PAY_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(P_WAGE_TYPE.EditValue) == string.Empty)
            {//급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                P_WAGE_TYPE_NAME.Focus();
                return;
            }

            if (IGR_FOOD_DED_PAYMENT.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            int vIDX_SELECT_YN = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("SELECT_YN");
            int vIDX_FOOD_DED_AMOUNT = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("FOOD_DED_AMOUNT");
            int vIDX_PERSON_ID = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("PERSON_ID");

            for (int vRow = 0; vRow < IGR_FOOD_DED_PAYMENT.RowCount; vRow++)
            {
                if (IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_FOOD_DED_PAYMENT.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);

                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_TRANS_STATUS", W_TRANS_YN.EditValue);
                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_PERSON_ID", IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_EXEC_FOOD_DED_PAYMENT.SetCommandParamValue("W_FOOD_DED_AMOUNT", IGR_FOOD_DED_PAYMENT.GetCellValue(vRow, vIDX_FOOD_DED_AMOUNT));
                    IDC_EXEC_FOOD_DED_PAYMENT.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_EXEC_FOOD_DED_PAYMENT.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_EXEC_FOOD_DED_PAYMENT.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_EXEC_FOOD_DED_PAYMENT.ExcuteError || mSTATUS == "F")
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
                    IGR_FOOD_DED_PAYMENT.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }

            IGR_FOOD_DED_PAYMENT.LastConfirmChanges();
            IDA_FOOD_DED_PAYMENT.OraSelectData.AcceptChanges();
            IDA_FOOD_DED_PAYMENT.Refillable = true;

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
            W_TRANS_YN.EditValue = vRadio.RadioCheckedString;
            if (iString.ISNull(W_TRANS_YN.EditValue) == "N")
            {
                BTN_SEND.Visible = true;
                BTN_CANCEL.Visible = false;
            }
            else
            {
                BTN_CANCEL.Visible = true;
                BTN_SEND.Visible = false;                
            } 
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_FOOD_DED_PAYMENT.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_YN = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("SELECT_YN");
            for (int vRow = 0; vRow < IGR_FOOD_DED_PAYMENT.RowCount; vRow++)
            {
                IGR_FOOD_DED_PAYMENT.SetCellValue(vRow, vIDX_SELECT_YN, CB_SELECT_YN.CheckBoxValue);
            }

            IGR_FOOD_DED_PAYMENT.LastConfirmChanges();
            IDA_FOOD_DED_PAYMENT.OraSelectData.AcceptChanges();
            IDA_FOOD_DED_PAYMENT.Refillable = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void IGR_YEAR_ADJUSTMENT__PAYMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_FOOD_DED_PAYMENT.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_FOOD_DED_PAYMENT.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_SELECT_YN)
            {
                IGR_FOOD_DED_PAYMENT.LastConfirmChanges();
                IDA_FOOD_DED_PAYMENT.OraSelectData.AcceptChanges();
                IDA_FOOD_DED_PAYMENT.Refillable = true;
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
         
        private void ilaYYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 2)));
        }
         
        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }
         
        #endregion

    }
}