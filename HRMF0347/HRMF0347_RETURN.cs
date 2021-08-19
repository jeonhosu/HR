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

namespace HRMF0347
{
    public partial class HRMF0347_RETURN : Office2007Form
    {        

        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCORP_ID = null;
        object mSTART_DATE = null;
        object mEND_DATE = null;
        object mAPPROVE_STATUS = null;
        object mFLOOR_ID = null;
        object mPERSON_ID = null;

        #endregion;

        #region ----- Constructor -----

        public HRMF0347_RETURN(ISAppInterface pAppInterface, object pCORP_ID, object pSTART_DATE, object pEND_DATE
                                , object pAPPROVE_STATUS, object pFLOOR_ID, object pPERSON_ID)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mCORP_ID = pCORP_ID;
            mSTART_DATE = pSTART_DATE;
            mEND_DATE = pEND_DATE;
            mAPPROVE_STATUS = pAPPROVE_STATUS;
            mFLOOR_ID = pFLOOR_ID;
            mPERSON_ID = pPERSON_ID;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iString.ISNull(mCORP_ID) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(mSTART_DATE) == string.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(mEND_DATE) == string.Empty)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (Convert.ToDateTime(mSTART_DATE) > Convert.ToDateTime(mEND_DATE))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_CORP_ID", mCORP_ID);
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_START_DATE", mSTART_DATE);
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_END_DATE", mEND_DATE);
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_APPROVE_STATUS", mAPPROVE_STATUS);
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_FLOOR_ID", mFLOOR_ID);
            idaHOLY_TYPE_RETURN.SetSelectParamValue("W_PERSON_ID", mPERSON_ID);
            idaHOLY_TYPE_RETURN.Fill();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
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

        private void HRMF0347_RETURN_Load(object sender, EventArgs e)
        {
            idaHOLY_TYPE_RETURN.FillSchema();
        }
        
        private void HRMF0347_RETURN_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        private void igrDUTY_PERIOD_CellDoubleClick(object pSender)
        {
            if (igrHOLY_TYPE.RowIndex < 0 && igrHOLY_TYPE.ColIndex == igrHOLY_TYPE.GetColumnToIndex("SELECT_YN"))
            {
                for (int r = 0; r < igrHOLY_TYPE.RowCount; r++)
                {
                    if (iString.ISNull(igrHOLY_TYPE.GetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("SELECT_YN")), "N") == "Y".ToString())
                    {
                        igrHOLY_TYPE.SetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("SELECT_YN"), "N");
                    }
                    else
                    {
                        igrHOLY_TYPE.SetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("SELECT_YN"), "Y");
                    }
                }
            }
        }

        private void ibtnSEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB();
        }

        private void ibtnSAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaHOLY_TYPE_RETURN.SetUpdateParamValue("P_APPROVE_STATUS", mAPPROVE_STATUS);
            idaHOLY_TYPE_RETURN.Update();

            this.Close();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaHOLY_TYPE_RETURN.Cancel();
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        
        #endregion

        #region ------ Adapter Event ------

        private void idaPERIOD_RETURN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["SELECT_YN"]) == "Y" && iString.ISNull(e.Row["REJECT_REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Reject Remark(반려사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SELECT_YN"]) == "N" && iString.ISNull(e.Row["REJECT_REMARK"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10276"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPERIOD_RETURN_UpdateCompleted(object pSender)
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "RETURN");
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "HOLY");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", mCORP_ID);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

    }
}