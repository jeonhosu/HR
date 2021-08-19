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

namespace HRMF0339
{
    public partial class HRMF0339_RETURN : Office2007Form
    {        

        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCORP_ID = null;
        DateTime mSYS_DATE = DateTime.Today;

        #endregion;

        #region ----- Constructor -----

        public HRMF0339_RETURN(ISAppInterface pAppInterface, object pCORP_ID, DateTime pSYS_DATE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mCORP_ID = pCORP_ID;
            mSYS_DATE = pSYS_DATE;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iConv.ISNull(mCORP_ID) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IGR_PERIOD_RETURN.LastConfirmChanges();
            IDA_WORKTIME_RETURN.OraSelectData.AcceptChanges();
            IDA_WORKTIME_RETURN.Refillable = true;

            IDA_WORKTIME_RETURN.SetSelectParamValue("W_CORP_ID", mCORP_ID);
            IDA_WORKTIME_RETURN.SetSelectParamValue("W_SYS_DATE", mSYS_DATE); 
            IDA_WORKTIME_RETURN.Fill();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private void Set_Update_Approve()
        {
            if (IGR_PERIOD_RETURN.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_PERIOD_RETURN.LastConfirmChanges();
            IDA_WORKTIME_RETURN.OraSelectData.AcceptChanges();
            IDA_WORKTIME_RETURN.Refillable = true;
            
            int vIDX_SELECT_YN = IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN");
            int vIDX_PERSON_ID = IGR_PERIOD_RETURN.GetColumnToIndex("PERSON_ID");
            int vIDX_WORK_DATE = IGR_PERIOD_RETURN.GetColumnToIndex("WORK_DATE");
            int vIDX_REJECT_REMARK = IGR_PERIOD_RETURN.GetColumnToIndex("REJECT_REMARK");
            int vIDX_APPROVE_STATUS = IGR_PERIOD_RETURN.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_PERIOD_RETURN.RowCount; i++)
            {
                if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                { 
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_WORK_DATE", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_WORK_DATE));
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_PERSON_ID", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_PERSON_ID));
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_CHECK_YN", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_REJECT_REMARK", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_REJECT_REMARK));
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_APPROVE_STATUS", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_SET_UPDATE_RETURN.SetCommandParamValue("P_SYS_DATE", mSYS_DATE);
                    IDC_SET_UPDATE_RETURN.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_UPDATE_RETURN.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_RETURN.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_UPDATE_RETURN.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            this.DialogResult = DialogResult.OK;
            this.Close();
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

        private void HRMF0339_RETURN_Load(object sender, EventArgs e)
        {
            IDA_WORKTIME_RETURN.FillSchema();
        }
        
        private void HRMF0339_RETURN_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
        }

        private void igrDUTY_PERIOD_CellDoubleClick(object pSender)
        {
            if (IGR_PERIOD_RETURN.RowIndex < 0 && IGR_PERIOD_RETURN.ColIndex == IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN"))
            {
                for (int r = 0; r < IGR_PERIOD_RETURN.RowCount; r++)
                {
                    if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(r, IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN")), "N") == "Y".ToString())
                    {
                        IGR_PERIOD_RETURN.SetCellValue(r, IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN"), "N");
                    }
                    else
                    {
                        IGR_PERIOD_RETURN.SetCellValue(r, IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN"), "Y");
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
            Set_Update_Approve();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_WORKTIME_RETURN.Cancel();
        }

        private void ibtnCLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----
        
        #endregion

        #region ------ Adapter Event ------

        private void idaPERIOD_RETURN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["SELECT_YN"]) == "Y" && iConv.ISNull(e.Row["REJECT_REMARK"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Reject Remark(반려사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SELECT_YN"]) == "N" && iConv.ISNull(e.Row["REJECT_REMARK"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10276"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPERIOD_RETURN_UpdateCompleted(object pSender)
        {
            IDC_GetDate.ExecuteNonQuery();
            object vLOCAL_DATE = iDate.ISGetDate(IDC_GetDate.GetCommandParamValue("X_LOCAL_DATE")).ToShortDateString();

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "RETURN");
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "DUTY");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", mCORP_ID);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

    }
}