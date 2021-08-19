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

namespace HRMF0634
{
    public partial class HRMF0634_CAL : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        string mYYYYMM; 
        string mCAL_TYPE;              // 급상여 구분.
        object mCORP_ID;
        object mDEPT_ID;
        object mFLOOR_ID;
        object mPERSON_ID; 
        string mALL_FLAG; 


        #endregion;

        #region ----- Constructor -----

        public HRMF0634_CAL(ISAppInterface pAppInterface, object pYYYYMM, object pCAL_TYPE, 
                            object pCORP_ID , object pCORP_NAME,
                            object pDEPT_ID , object pDEPT_NAME,
                            object pFLOOR_ID, object pFLOOR_NAME,
                            object pALL_FLAG, 
                            object pPERSON_ID, object pNAME)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mCAL_TYPE = iString.ISNull(pCAL_TYPE);
            mYYYYMM  = iString.ISNull(pYYYYMM);
            mCORP_ID = pCORP_ID;// iString.ISDecimaltoZero(pCORP_ID);
            mDEPT_ID = pDEPT_ID;// iString.ISDecimaltoZero(pDEPT_ID);
            mFLOOR_ID = pFLOOR_ID;// iString.ISDecimaltoZero(pFLOOR_ID);
            mPERSON_ID = pPERSON_ID;
            mALL_FLAG = iString.ISNull(pALL_FLAG);

            V_DEPT_NAME.EditValue = pDEPT_NAME;
            V_FLOOR_NAME.EditValue = pFLOOR_NAME;
            V_NAME.EditValue = pNAME;
        }

        #endregion;

        #region ----- Private Methods ----
          
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

        #region ----- Form Event ------

        private void HRMF0634_CAL_Load(object sender, EventArgs e)
        {
             
        }

        private void HRMF0634_CAL_Shown(object sender, EventArgs e)
        {
            W_RESERVE_YYYYMM.EditValue = mYYYYMM;
            CORP_ID.EditValue = mCORP_ID;// iString.ISDecimaltoZero(mCORP_ID);
            DEPT_ID.EditValue = mDEPT_ID; // iString.ISDecimaltoZero(mDEPT_ID); 
            FLOOR_ID.EditValue = mFLOOR_ID; // iString.ISDecimaltoZero(mFLOOR_ID); 
            ALL_FLAG.EditValue = mALL_FLAG;
            PERSON_ID.EditValue = mPERSON_ID; // iString.ISDecimaltoZero(mPERSON_ID);  
           if(mCAL_TYPE == "CAL")
            {
                itb_CLOSED.Visible = false;
                itb_CLOSED_CANCEL.Visible = false;
                itb_CAL.Visible = true; 
            }
           else
            {
                itb_CLOSED.Visible = true;
                itb_CLOSED_CANCEL.Visible = true;
                itb_CAL.Visible = false; 
            } 
        }
        
        private void btnSAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
       
        }

        // 창닫기
        private void btnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void igrPAY_PERIOD_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {

        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPAY_PERIOD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //if (iString.ISNull(e.Row["ADJUSTMENT_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["OLD_PAY_YYYYMM"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["START_DATE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["END_DATE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        #endregion

        private void ILA_YYYYMM_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_STD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 3)));
        }

        private void itb_CAL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //계산
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }
           
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            // 실행.
            string mStatus = "F";
            string mMessage = null; 
            // idcRETIRE_CALCULATE.SetCommandParamValue("W_RETIRE_CAL_TYPE", pRETIRE_CAL_TYPE);
            idcRESERVE_DC_CAL.ExecuteNonQuery();
            mStatus = iString.ISNull(idcRESERVE_DC_CAL.GetCommandParamValue("O_STATUS"));
            mMessage = iString.ISNull(idcRESERVE_DC_CAL.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (idcRESERVE_DC_CAL.ExcuteError || mStatus == "F")
            { 
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }  
            this.Close(); 
        }

        private void itb_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //계산
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty; 

             
            IDC_SET_RETIRE_RESERVE_CLOSED.SetCommandParamValue("P_CLOSED_FLAG", "Y");
            IDC_SET_RETIRE_RESERVE_CLOSED.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_SET_RETIRE_RESERVE_CLOSED.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_SET_RETIRE_RESERVE_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_RETIRE_RESERVE_CLOSED.ExcuteError || mSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (mMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }  

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            this.Close();
        }

        private void itb_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //계산
            if (iString.ISNull(W_RESERVE_YYYYMM.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_RESERVE_YYYYMM.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string mSTATUS = "F";
            string mMESSAGE = String.Empty;

            IDC_SET_RETIRE_RESERVE_CLOSED.SetCommandParamValue("P_CLOSED_FLAG", "N");
            IDC_SET_RETIRE_RESERVE_CLOSED.ExecuteNonQuery();
            mSTATUS = iString.ISNull(IDC_SET_RETIRE_RESERVE_CLOSED.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(IDC_SET_RETIRE_RESERVE_CLOSED.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_RETIRE_RESERVE_CLOSED.ExcuteError || mSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (mMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            } 

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            this.Close();
        }
        
    }
}