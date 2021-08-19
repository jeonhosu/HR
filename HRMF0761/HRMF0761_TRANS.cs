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

namespace HRMF0761
{
    public partial class HRMF0761_TRANS : Office2007Form
    {
       
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0761_TRANS(Form pMainForm, ISAppInterface pAppInterface
                                , string pExec_Type 
                                , object pYYYY
                                , object pCORP_NAME, object pCORP_ID
                                , object pOPERATING_UNIT_NAME, object pOPERATING_UNIT_ID 
                                , object pDEPT_NAME, object pDEPT_ID
                                , object pFLOOR_NAME, object pFLOOR_ID
                                , object pYEAR_EMPLOYE_TYPE_DESC, object pYEAR_EMPLOYE_TYPE 
                                , object pNAME, object pPERSON_NUM, object pPERSON_ID)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_CORP_ID.EditValue = pCORP_ID;
            W_CORP_NAME.EditValue = pCORP_NAME;
            W_OPERATING_UNIT_ID.EditValue = pOPERATING_UNIT_ID;
            W_OPERATING_UNIT_NAME.EditValue = pOPERATING_UNIT_NAME;
            W_DEPT_ID.EditValue = pDEPT_ID;
            W_DEPT_NAME.EditValue = pDEPT_NAME;
            W_FLOOR_ID.EditValue = pFLOOR_ID;
            W_FLOOR_DESC.EditValue = pFLOOR_NAME;
            W_YEAR_EMPLOYE_TYPE.EditValue = pYEAR_EMPLOYE_TYPE;
            W_YEAR_EMPLOYE_TYPE_DESC.EditValue = pYEAR_EMPLOYE_TYPE_DESC;
            W_PERSON_ID.EditValue = pPERSON_ID;
            W_PERSON_NUM.EditValue = pPERSON_NUM;
            W_NAME.EditValue = pNAME;

            W_YEAR_YYYYMM.EditValue = pYYYY;

            if (pExec_Type.ToUpper() == "EXEC")
            {
                W_RB_NO.CheckedState = ISUtil.Enum.CheckedState.Checked;
                W_TRANS_YN.EditValue = W_RB_NO.RadioCheckedString;
                BTN_SEND.Visible = true;
                BTN_CANCEL.Visible = false;
            }
            else
            {
                W_RB_YES.CheckedState = ISUtil.Enum.CheckedState.Checked;
                W_TRANS_YN.EditValue = W_RB_YES.RadioCheckedString;
                BTN_CANCEL.Visible = true;
                BTN_SEND.Visible = false;
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
            if (W_YEAR_YYYYMM.EditValue == null)
            {//적용 년도  FCM_10036
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10068"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYYMM.Focus();
                return;
            }

            IDA_TRAN_PERSON_LIST.Fill();
            IGR_TRAN_PERSON_LIST.Focus();
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

        private void HRMF0761_TRANS_Load(object sender, EventArgs e)
        {
           
        }

        private void HRMF0761_TRANS_Shown(object sender, EventArgs e)
        {
             
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void BTN_SEND_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 연말정산내역 급여전송 내용                        
            if (iString.ISNull(W_YEAR_YYYYMM.EditValue) == String.Empty)
            {// 적용 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYYMM.Focus();
                return;
            } 

            if (IGR_TRAN_PERSON_LIST.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            int vIDX_SELECT_YN = IGR_TRAN_PERSON_LIST.GetColumnToIndex("SELECT_YN");
            int vIDX_CORP_ID = IGR_TRAN_PERSON_LIST.GetColumnToIndex("CORP_ID");
            int vIDX_ADJUST_YYYY = IGR_TRAN_PERSON_LIST.GetColumnToIndex("ADJUST_YYYY");
            int vIDX_PERSON_ID = IGR_TRAN_PERSON_LIST.GetColumnToIndex("PERSON_ID");
            int vIDX_PERSON_NUM = IGR_TRAN_PERSON_LIST.GetColumnToIndex("PERSON_NUM");
            int vIDX_NAME = IGR_TRAN_PERSON_LIST.GetColumnToIndex("NAME");
             
            for (int vRow = 0; vRow < IGR_TRAN_PERSON_LIST.RowCount; vRow++)
            {
                if (IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_TRAN_PERSON_LIST.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);

                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_SELECT_YN", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_SELECT_YN));
                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_CORP_ID", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_CORP_ID));
                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_YYYY", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_ADJUST_YYYY));
                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_PERSON_ID", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_PERSON_NUM", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_PERSON_NUM));
                    IDC_EXEC_YEAR_ADJUST.SetCommandParamValue("P_NAME", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_NAME));
                    IDC_EXEC_YEAR_ADJUST.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_EXEC_YEAR_ADJUST.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_EXEC_YEAR_ADJUST.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_EXEC_YEAR_ADJUST.ExcuteError || mSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                           
                        }
                        return;
                    }
                    IGR_TRAN_PERSON_LIST.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }

            IGR_TRAN_PERSON_LIST.LastConfirmChanges();
            IDA_TRAN_PERSON_LIST.OraSelectData.AcceptChanges();
            IDA_TRAN_PERSON_LIST.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            SEARCH_DB();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(W_YEAR_YYYYMM.EditValue) == String.Empty)
            {// 적용 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YEAR_YYYYMM.Focus();
                return;
            } 
            
            if (IGR_TRAN_PERSON_LIST.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            int vIDX_SELECT_YN = IGR_TRAN_PERSON_LIST.GetColumnToIndex("SELECT_YN");
            int vIDX_CORP_ID = IGR_TRAN_PERSON_LIST.GetColumnToIndex("CORP_ID");
            int vIDX_ADJUST_YYYY = IGR_TRAN_PERSON_LIST.GetColumnToIndex("ADJUST_YYYY");
            int vIDX_PERSON_ID = IGR_TRAN_PERSON_LIST.GetColumnToIndex("PERSON_ID");
            int vIDX_PERSON_NUM = IGR_TRAN_PERSON_LIST.GetColumnToIndex("PERSON_NUM");
            int vIDX_NAME = IGR_TRAN_PERSON_LIST.GetColumnToIndex("NAME");
            
            for (int vRow = 0; vRow < IGR_TRAN_PERSON_LIST.RowCount; vRow++)
            {
                if (IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_TRAN_PERSON_LIST.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);

                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_SELECT_YN", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_SELECT_YN));
                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_CORP_ID", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_CORP_ID));
                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_YYYY", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_ADJUST_YYYY));
                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_PERSON_ID", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_PERSON_NUM", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_PERSON_NUM));
                    IDC_CANCEL_YEAR_ADJUST.SetCommandParamValue("P_NAME", IGR_TRAN_PERSON_LIST.GetCellValue(vRow, vIDX_NAME));

                    IDC_CANCEL_YEAR_ADJUST.ExecuteNonQuery();
                    mSTATUS = iString.ISNull(IDC_CANCEL_YEAR_ADJUST.GetCommandParamValue("O_STATUS"));
                    mMESSAGE = iString.ISNull(IDC_CANCEL_YEAR_ADJUST.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_EXEC_YEAR_ADJUST.ExcuteError || mSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        if (mMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                            
                        }
                        return;
                    }
                    IGR_TRAN_PERSON_LIST.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }

            IGR_TRAN_PERSON_LIST.LastConfirmChanges();
            IDA_TRAN_PERSON_LIST.OraSelectData.AcceptChanges();
            IDA_TRAN_PERSON_LIST.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
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
            if (IGR_TRAN_PERSON_LIST.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_SELECT_YN = IGR_TRAN_PERSON_LIST.GetColumnToIndex("SELECT_YN");
            for (int vRow = 0; vRow < IGR_TRAN_PERSON_LIST.RowCount; vRow++)
            {
                IGR_TRAN_PERSON_LIST.SetCellValue(vRow, vIDX_SELECT_YN, CB_SELECT_YN.CheckBoxValue);
            }

            IGR_TRAN_PERSON_LIST.LastConfirmChanges();
            IDA_TRAN_PERSON_LIST.OraSelectData.AcceptChanges();
            IDA_TRAN_PERSON_LIST.Refillable = true;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void IGR_YEAR_ADJUSTMENT__PAYMENT_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_TRAN_PERSON_LIST.RowCount < 1)
            {
                return;
            }

            int vIDX_SELECT_YN = IGR_TRAN_PERSON_LIST.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_SELECT_YN)
            {
                IGR_TRAN_PERSON_LIST.LastConfirmChanges();
                IDA_TRAN_PERSON_LIST.OraSelectData.AcceptChanges();
                IDA_TRAN_PERSON_LIST.Refillable = true;
            }
        }

        #endregion              

        #region ----- Lookup Event -----
         
        private void ILA_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
          
        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        private void ILA_PERSON_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_YYYYMM", string.Format("{0}-12", W_YEAR_YYYYMM.EditValue));
        }

    }
}