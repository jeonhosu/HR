using System;
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

namespace HRMF0337
{
    public partial class HRMF0337 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private bool mIsFirst = false;

        #endregion;

        #region ----- Constructor -----

        public HRMF0337(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            //if (START_DATE_0.EditValue == null)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    START_DATE_0.Focus();
            //    return;
            //}

            //if (END_DATE_0.EditValue == null)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    END_DATE_0.Focus();
            //    return;
            //}

            if (WORK_DATE_FR.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_FR.Focus();
                return;
            }

            if (WORK_DATE_TO.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_TO.Focus();
                return;
            }

            idaOT_HEADER.OraSelectData.AcceptChanges();
            idaOT_HEADER.Refillable = true;

            idaOT_HEADER.Fill();
            igrOT_HEADER.Focus();
        }

        #endregion;

        #region ----- MDi ToolBar Button Events -----

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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaOT_HEADER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaOT_HEADER.IsFocused)
                    {
                        idaOT_HEADER.Cancel();
                    }
                    else if (idaOT_LINE.IsFocused)
                    {
                        idaOT_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaOT_HEADER.IsFocused)
                    {
                        idaOT_HEADER.Delete();
                    }
                    else if (idaOT_LINE.IsFocused)
                    {
                        idaOT_LINE.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0337_Load(object sender, EventArgs e)
        {
            //REQUEST_DATE_FR_0.EditValue = DateTime.Today.AddDays(-7);
            //REQUEST_DATE_TO_0.EditValue = DateTime.Today.AddDays(3);

            WORK_DATE_FR.EditValue = DateTime.Today.AddDays(-7);
            WORK_DATE_TO.EditValue = DateTime.Today.AddDays(3);

            //WORK_DATE_FR.EditValue = new System.DateTime(2011, 8, 1);
            //WORK_DATE_TO.EditValue = new System.DateTime(2011, 8, 31);

            idaOT_HEADER.FillSchema();
            DefaultCorporation();

            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EMAIL_STATUS.EditValue = "N";
        }

        private void HRMF0337_Shown(object sender, EventArgs e)
        {
            mIsFirst = true;
        }

        private void ibtOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 승인
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            idaOT_HEADER.SetUpdateParamValue("P_APPROVE_FLAG", "OK");
            idaOT_HEADER.SetUpdateParamValue("P_APPROVE_STATUS", APPROVE_STATUS_0.EditValue);
            idaOT_HEADER.Update();

            SEARCH_DB();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 취소
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            idaOT_HEADER.SetUpdateParamValue("P_APPROVE_FLAG", "CANCEL");
            idaOT_HEADER.SetUpdateParamValue("P_APPROVE_STATUS", APPROVE_STATUS_0.EditValue);
            idaOT_HEADER.Update();

            SEARCH_DB();
        }

        private void irbSTATUS_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            APPROVE_STATUS_0.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();

            if (mIsFirst == false)
            {
                return;
            }

            SEARCH_DB();
        }

        private void igrOT_HEADER_CellDoubleClick(object pSender)
        {
            string mAPPROVE_STATE = iString.ISNull(APPROVE_STATUS_0.EditValue);
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                return;
            }

            if (igrOT_HEADER.RowIndex < 0 && igrOT_HEADER.ColIndex == igrOT_HEADER.GetColumnToIndex("APPROVE_YN"))
            {
                for (int r = 0; r < igrOT_HEADER.RowCount; r++)
                {
                    if (iString.ISNull(igrOT_HEADER.GetCellValue(r, igrOT_HEADER.GetColumnToIndex("APPROVE_YN")), "N") == "Y".ToString())
                    {
                        igrOT_HEADER.SetCellValue(r, igrOT_HEADER.GetColumnToIndex("APPROVE_YN"), "N");                        
                    }
                    else
                    {
                        igrOT_HEADER.SetCellValue(r, igrOT_HEADER.GetColumnToIndex("APPROVE_YN"), "Y");
                    }
                }
            }
        }

        #endregion

        #region ----- Adapter Event ------

        private void idaOT_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Row["REQ_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Type(신청구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["REQ_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Date(신청 일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DUTY_MANAGER_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Control Level(근태관리 단위)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["REQ_PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Person(신청자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaOT_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (igrOT_LINE.RowCount != 0)
            {// 라인 존재.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (e.Row["OT_HEADER_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Number(신청 번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaOT_HEADER_UpdateCompleted(object pSender)
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "OT");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

        #region ----- LookUP Event ----

        private void ilaDUTY_MANAGER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDUTY_MANAGER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
            ildDUTY_MANAGER.SetLookupParamValue("W_CAP_CHECK_YN", "Y");
        }
        
        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", WORK_DATE_TO.EditValue);
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", WORK_DATE_TO.EditValue);
        }

        #endregion

        #region ----- Button Event -----

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iString.ISNull(APPROVE_STATUS_0.EditValue);
            int mIDX_SELECT_YN = igrOT_HEADER.GetColumnToIndex("APPROVE_YN");
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                btnOK.Enabled = false;
                btnCANCEL.Enabled = false;
                btnRETURN.Enabled = false;

                igrOT_HEADER.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                btnOK.Enabled = true;
                btnCANCEL.Enabled = true;
                btnRETURN.Enabled = true;

                igrOT_HEADER.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
        }

        private void btnRETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_FR.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_FR.Focus();
                return;
            }
            if (WORK_DATE_TO.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_TO.Focus();
                return;
            }
            if (Convert.ToDateTime(WORK_DATE_FR.EditValue) > Convert.ToDateTime(WORK_DATE_TO.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_TO.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            DialogResult dlgResultValue;
            HRMF0337_RETURN vHRMF0307_RETURN = new HRMF0337_RETURN(isAppInterfaceAdv1.AppInterface
                                                                  , CORP_ID_0.EditValue
                                                                  , WORK_DATE_FR.EditValue
                                                                  , WORK_DATE_TO.EditValue
                                                                  , APPROVE_STATUS_0.EditValue
                                                                  , OT_HEADER_ID_0.EditValue
                                                                  , DUTY_MANAGER_ID_0.EditValue
                                                                  );
            dlgResultValue = vHRMF0307_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            vHRMF0307_RETURN.Dispose();

            SEARCH_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }

        #endregion

        #region ----- Edit Event -----

        private void WORK_DATE_FR_EditValueChanged(object pSender)
        {
            System.DateTime vDate = WORK_DATE_FR.DateTimeValue;
            WORK_DATE_TO.EditValue = vDate.AddDays(10);
        }

        #endregion
    }
}