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
    public partial class HRMF0347 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----

        public HRMF0347(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime( END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            idaHOLY_TYPE.OraSelectData.AcceptChanges();
            idaHOLY_TYPE.Refillable = true;

            idaHOLY_TYPE.SetSelectParamValue("W_SEARCH_TYPE", "R");
            idaHOLY_TYPE.Fill();
            igrHOLY_TYPE.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pStart_Date, Object pEnd_Date)
        {            
            idaWORK_CALENDAR.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            idaWORK_CALENDAR.SetSelectParamValue("W_START_DATE", pStart_Date);
            idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pEnd_Date);

            if (pStart_Date != DBNull.Value && pEnd_Date != DBNull.Value)
            {
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_START_YEAR", iDate.ISYear(Convert.ToDateTime(pStart_Date)));
                idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_END_YEAR", iDate.ISYear(Convert.ToDateTime(pEnd_Date)));
            }
            idaWORK_CALENDAR.Fill();
            idaHOLIDAY_MANAGEMENT.Fill();
        }

        private bool isAdd_DB_Check()
        {// 데이터 추가시 검증.
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return false;
            }
            return true;
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE = iString.ISNull(APPROVE_STATUS_0.EditValue);
            int mIDX_SELECT_YN = igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN");
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                btnOK.Enabled = false;
                btnCANCEL.Enabled = false;
                btnRETURN.Enabled = false;

                igrHOLY_TYPE.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                btnOK.Enabled = true;
                btnCANCEL.Enabled = true;
                btnRETURN.Enabled = true;

                igrHOLY_TYPE.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            }
        }

        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (igrHOLY_TYPE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_APPROVE_YN = igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN");
            int vIDX_HOLY_TYPE_ID = igrHOLY_TYPE.GetColumnToIndex("HOLY_TYPE_ID");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < igrHOLY_TYPE.RowCount; i++)
            {
                if (iString.ISNull(igrHOLY_TYPE.GetCellValue(i, vIDX_APPROVE_YN), "N") == "Y")
                {

                    IDC_UPDATE_APPROVE.SetCommandParamValue("W_HOLY_TYPE_ID", igrHOLY_TYPE.GetCellValue(i, vIDX_HOLY_TYPE_ID));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", igrHOLY_TYPE.GetCellValue(i, vIDX_APPROVE_YN));
                    IDC_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_FLAG", pApproved_Flag);
                    IDC_UPDATE_APPROVE.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDC_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDC_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
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

            // eMail 전송.
            Send_Mail();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void Send_Mail()
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "HOLY");
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery(); 
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----
        
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
                    if(idaHOLY_TYPE.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {                
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        idaHOLY_TYPE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHOLY_TYPE.IsFocused)
                    {
                        idaHOLY_TYPE.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0347_Load(object sender, EventArgs e)
        {
            idaHOLY_TYPE.FillSchema();
            START_DATE_0.EditValue = DateTime.Today.AddDays(-7);
            END_DATE_0.EditValue = DateTime.Today.AddDays(7);
            
            // CORP SETTING
            DefaultCorporation();
            
            //LOOKUP SETTING
            ildAPPROVE_STATUS.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ildAPPROVE_STATUS.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EMAIL_STATUS.EditValue = "N";
        }

        private void btnOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 승인
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            Set_Update_Approve("OK");
        }

        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 취소
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_CANCEL";
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
            Set_Update_Approve("CANCEL");
        }

        private void btnRETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime(END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            DialogResult dlgResultValue;
            Form vHRMF0347_RETURN = new HRMF0347_RETURN(isAppInterfaceAdv1.AppInterface
                                                        , CORP_ID_0.EditValue
                                                        , START_DATE_0.EditValue
                                                        , END_DATE_0.EditValue
                                                        , APPROVE_STATUS_0.EditValue
                                                        , FLOOR_ID_0.EditValue
                                                        , PERSON_ID_0.EditValue
                                                        );
            dlgResultValue = vHRMF0347_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            vHRMF0347_RETURN.Dispose();

            Search_DB();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private void irbSTATUS_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            APPROVE_STATUS_0.EditValue = iStatus.RadioCheckedString;
            Set_BTN_STATE();  // 버튼 상태 변경.
            Search_DB();
        }

        private void igrHOLY_TYPE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_APPROVE_YN = igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN");
            if (e.ColIndex == vIDX_APPROVE_YN)
            {
                igrHOLY_TYPE.LastConfirmChanges();
                idaHOLY_TYPE.OraSelectData.AcceptChanges();
                idaHOLY_TYPE.Refillable = true;
            }
        }

        private void igrHOLY_TYPE_CellDoubleClick(object pSender)
        {
            string mAPPROVE_STATE = iString.ISNull(APPROVE_STATUS_0.EditValue);
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                return;
            }
            if (igrHOLY_TYPE.RowIndex < 0 && igrHOLY_TYPE.ColIndex == igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN"))
            {
                for (int r = 0; r < igrHOLY_TYPE.RowCount; r++)
                {
                    if (iString.ISNull(igrHOLY_TYPE.GetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN")), "N") == "Y".ToString())
                    {
                        igrHOLY_TYPE.SetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN"), "N");
                    }
                    else
                    {
                        igrHOLY_TYPE.SetCellValue(r, igrHOLY_TYPE.GetColumnToIndex("APPROVE_YN"), "Y");
                    }
                }
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDUTY_PERIOD_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if(e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=사원 정보"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["START_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=시작일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["END_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=종료일자"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Convert.ToDateTime(e.Row["START_DATE"]) > Convert.ToDateTime(e.Row["END_DATE"]))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDUTY_PERIOD_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row["APPROVE_STATUS"].ToString() != "A".ToString() ||
                e.Row["APPROVE_STATUS"].ToString() != "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=이미 승인된 자료는"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDUTY_PERIOD_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            //if (igrDUTY_PERIOD.RowCount != 0)
            //{
                isSearch_WorkCalendar(igrHOLY_TYPE.GetCellValue("PERSON_ID"), igrHOLY_TYPE.GetCellValue("START_DATE"), igrHOLY_TYPE.GetCellValue("END_DATE"));
            //}
        }

        private void idaHOLY_TYPE_UpdateCompleted(object pSender)
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "HOLY");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

        #region ----- LookUp Event -----
        private void ilaAPPROVE_STATUS_0_SelectedRowData(object pSender)
        {
            idaHOLY_TYPE.Fill();
        }


        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaHOLY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        #endregion

    }
}