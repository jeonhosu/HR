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

namespace HRMF0603
{
    public partial class HRMF0603_PAYMENT_PERIOD : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        int mADJUSTMENT_ID;             //퇴직정산 ID
        string mWAGE_TYPE;              // 급상여 구분.

        #endregion;

        #region ----- Constructor -----

        public HRMF0603_PAYMENT_PERIOD(ISAppInterface pAppInterface, object pADJUSTMENT_ID, object pWAGE_TYPE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mADJUSTMENT_ID = iString.ISNumtoZero(pADJUSTMENT_ID);
            mWAGE_TYPE = iString.ISNull(pWAGE_TYPE);
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            idaPAY_PERIOD.SetSelectParamValue("W_ADJUSTMENT_ID", mADJUSTMENT_ID);
            idaPAY_PERIOD.SetSelectParamValue("W_WAGE_TYPE", mWAGE_TYPE);
            idaPAY_PERIOD.Fill();
        }

        private void INIT_DAY_COUNT(object pSTART_DATE, object pEND_DATE)
        {
            object mDAY_COUNT;

            idcPERIOD_DAY.SetCommandParamValue("P_START_DATE", pSTART_DATE);
            idcPERIOD_DAY.SetCommandParamValue("P_END_DATE", pEND_DATE);
            idcPERIOD_DAY.SetCommandParamValue("P_ADD_DAY", 1);
            idcPERIOD_DAY.ExecuteNonQuery();
            mDAY_COUNT = idcPERIOD_DAY.GetCommandParamValue("O_DAY_COUNT");
            igrPAY_PERIOD.SetCellValue("DAY_COUNT", mDAY_COUNT);
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

        #region ----- Form Event ------

        private void HRMF0603_PAYMENT_PERIOD_Load(object sender, EventArgs e)
        {
            idaPAY_PERIOD.FillSchema();
        }

        private void HRMF0603_PAYMENT_PERIOD_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        // 명세서 발급
        private void btnSAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            btnSAVE.Focus();
            idaPAY_PERIOD.Update();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // 명세서 발급 취소
        private void btnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void igrPAY_PERIOD_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            if (e.ColIndex == igrPAY_PERIOD.GetColumnToIndex("START_DATE"))
            {
                INIT_DAY_COUNT(e.CellValue, igrPAY_PERIOD.GetCellValue("END_DATE"));
            }
            else if (e.ColIndex == igrPAY_PERIOD.GetColumnToIndex("END_DATE"))
            {
                INIT_DAY_COUNT(igrPAY_PERIOD.GetCellValue("START_DATE"), e.CellValue);
            }
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPAY_PERIOD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ADJUSTMENT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["OLD_PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["START_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["END_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion
    }
}