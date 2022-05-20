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

namespace HRMF0251
{
    public partial class HRMF0251_RETURN : Office2007Form
    {        

        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
         
        DateTime m_Start_DATE = DateTime.Today;
        DateTime m_End_DATE = DateTime.Today;

        #endregion;

        #region ----- Constructor -----

        public HRMF0251_RETURN(ISAppInterface pAppInterface, object pCORP_ID, DateTime pStart_Date, DateTime pEnd_Date)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_CORP_ID.EditValue = pCORP_ID;
            m_Start_DATE = pStart_Date;
            m_End_DATE = pEnd_Date;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (iConv.ISNull(V_CORP_ID.EditValue) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IGR_PERIOD_RETURN.LastConfirmChanges();
            IDA_PERIOD_RETURN.OraSelectData.AcceptChanges();
            IDA_PERIOD_RETURN.Refillable = true;
             
            IDA_PERIOD_RETURN.Fill();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();
        }

        private void Set_Update_Approve()
        {
            if (IGR_PERIOD_RETURN.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            IGR_PERIOD_RETURN.LastConfirmChanges();
            IDA_PERIOD_RETURN.OraSelectData.AcceptChanges();
            IDA_PERIOD_RETURN.Refillable = true;
            
            int vIDX_SELECT_YN = IGR_PERIOD_RETURN.GetColumnToIndex("SELECT_YN");
            int vIDX_PRINT_REQ_NUM = IGR_PERIOD_RETURN.GetColumnToIndex("PRINT_REQ_NUM");
            int vIDX_REJECT_REMARK = IGR_PERIOD_RETURN.GetColumnToIndex("REJECT_REMARK");

            for (int i = 0; i < IGR_PERIOD_RETURN.RowCount; i++)
            {
                if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                {
                    if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN)) == "Y" && iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_REJECT_REMARK)) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", "&&FIELD_NAME:=Reject Remark(반려사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN)) == "N" && iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_REJECT_REMARK)) != string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10276"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    } 
                }
            }

            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_PERIOD_RETURN.RowCount; i++)
            {
                if (iConv.ISNull(IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                { 
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_PRINT_REQ_NUM", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_PRINT_REQ_NUM));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_REJECT_REMARK", IGR_PERIOD_RETURN.GetCellValue(i, vIDX_REJECT_REMARK));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_APPROVE_STATUS", "R"); 
                    IDC_SET_UPDATE_APPROVE.ExecuteNonQuery();
                     
                    vSTATUS = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_UPDATE_APPROVE.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        MessageBoxAdv.Show(IDC_SET_UPDATE_APPROVE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
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

        private void HRMF0251_RETURN_Load(object sender, EventArgs e)
        {
            IDA_PERIOD_RETURN.FillSchema();
        }
        
        private void HRMF0251_RETURN_Shown(object sender, EventArgs e)
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
            IDA_PERIOD_RETURN.Cancel();
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
         

        #endregion

    }
}