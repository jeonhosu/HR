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
    public partial class HRMF0339 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime(); 
        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #endregion;

        #region ----- Constructor -----

        public HRMF0339(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildWORK_CORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildWORK_CORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            WORK_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            WORK_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void CAP_Status()
        {
            //작업장 
            IDC_DEFAULT_CAP.SetCommandParamValue("W_MODULE_CODE", "20");
            IDC_DEFAULT_CAP.SetCommandParamValue("W_END_DATE", W_WORK_DATE_FR.EditValue);
            IDC_DEFAULT_CAP.ExecuteNonQuery();
            CAP_STATUS.EditValue = IDC_DEFAULT_CAP.GetCommandParamValue("O_CAP_LEVEL");
            
        }
        

        private void DefaultFloor()
        {
            //작업장
            idcDEFAULT_FLOOR.ExecuteNonQuery();
            FLOOR_NAME_0.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            FLOOR_ID_0.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");
        }

        private void Search_DB()
        {
            if (WORK_CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_CORP_NAME_0.Focus();
                return;
            }
            if (W_WORK_DATE_FR.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_FR.Focus();
                return;
            }

            CHECK.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            IDA_DAY_WORK_MODIFY_APPR.Cancel();
            IDA_DAY_WORK_MODIFY_APPR.Fill();
            IGR_DAY_WORK_MODIFY_APPR.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = pWork_Date;

            idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pWork_Date);
            idaDAY_HISTORY.Fill();
            idaDUTY_PERIOD.Fill();
            idaWORK_CALENDAR.Fill();
        }

        private void isSearch_Day_History(int pAdd_Day)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(WORK_DATE_8.EditValue) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = Convert.ToDateTime(WORK_DATE_8.EditValue).AddDays(pAdd_Day);
            idaDAY_HISTORY.Fill();
        }

        #endregion;

        #region ----- MDi ToolBar Button Events -----

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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_DAY_WORK_MODIFY_APPR.IsFocused)
                    {
                        IDA_DAY_WORK_MODIFY_APPR.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DAY_WORK_MODIFY_APPR.IsFocused)
                    {
                        IDA_DAY_WORK_MODIFY_APPR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DAY_WORK_MODIFY_APPR.IsFocused)
                    {
                        IDA_DAY_WORK_MODIFY_APPR.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0339_Load(object sender, EventArgs e)
        {
            W_WORK_DATE_FR.EditValue = DateTime.Today.AddDays(-31); 
            W_WORK_DATE_TO.EditValue = DateTime.Today; 

            DefaultCorporation();
            //DefaultFloor();
            CAP_Status();

            WORK_CORP_NAME_0.BringToFront();
            BTN_OK.BringToFront();
            RB_N.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_APPROVE_STATUS.EditValue = RB_N.RadioCheckedString;
            EMAIL_STATUS.EditValue = "N";

            IDA_DAY_WORK_MODIFY_APPR.FillSchema(); 
        }

        private void HRMF0339_Shown(object sender, EventArgs e)
        {
            Set_BTN_STATE();
        }

        private void CHECK_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN");
            for (int r = 0; r < IGR_DAY_WORK_MODIFY_APPR.RowCount; r++)
            {
                IGR_DAY_WORK_MODIFY_APPR.SetCellValue(r, vIDX_SELECT_FLAG, CHECK.CheckBoxString);
            }
        }

        private void igrDAY_INTERFACE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_APPROVE_YN = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_APPROVE_YN)
            {
                IGR_DAY_WORK_MODIFY_APPR.LastConfirmChanges();
                IDA_DAY_WORK_MODIFY_APPR.OraSelectData.AcceptChanges();
                IDA_DAY_WORK_MODIFY_APPR.Refillable = true;
            }
        }

        #endregion;

        #region ----- Radio Button Event -----

        private void RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_APPROVE_STATUS.EditValue = iStatus.RadioCheckedString;

            Set_BTN_STATE();  // 버튼 상태 변경.
            Search_DB();
        }

        private void Set_BTN_STATE()
        {
            string mAPPROVE_STATE =  iConv.ISNull(V_APPROVE_STATUS.EditValue);
            int mIDX_SELECT_YN = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN");
            if (mAPPROVE_STATE == String.Empty || mAPPROVE_STATE == "R")
            {
                BTN_OK.Enabled = false;
                BTN_CANCEL.Enabled = false;
                BTN_RETURN.Enabled = false;

                IGR_DAY_WORK_MODIFY_APPR.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 0;
            }
            else
            {
                if (mAPPROVE_STATE == "N")
                {
                    BTN_OK.Enabled = true;
                    BTN_CANCEL.Enabled = false;
                }
                else
                {
                    BTN_OK.Enabled = false;
                    BTN_CANCEL.Enabled = true;
                }
                BTN_RETURN.Enabled = true;
                IGR_DAY_WORK_MODIFY_APPR.GridAdvExColElement[mIDX_SELECT_YN].Updatable = 1;
            } 
        }

        #endregion;

        #region ----- Button Event -----

        private void ibtOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //string vAPPROVE_STATUS = null;
            // 승인
            // EMAIL STATUS.
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A2".ToString())
            {
                EMAIL_STATUS.EditValue = "A2_OK";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }

            Set_Update_Approve("OK"); 
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 취소
            if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A1".ToString())
            {
                EMAIL_STATUS.EditValue = "A1_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "A2".ToString())
            {
                EMAIL_STATUS.EditValue = "A2_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iConv.ISNull(V_APPROVE_STATUS.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            Set_Update_Approve("CANCEL");

            //// refill.
            //idaDAY_INTERFACE.Cancel();
            //Search_DB();
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }

        //반려
        private void btnRETURN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (WORK_CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_CORP_NAME_0.Focus();
                return;
            }
            if (W_WORK_DATE_FR.EditValue == null)
            {// 작업일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WORK_DATE_FR.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            DialogResult dlgResultValue;

            //작업일자 
            IDC_GET_LOCAL_DATETIME_P.ExecuteNonQuery();
            DateTime vLOCAL_DATE = iDate.ISGetDate(IDC_GET_LOCAL_DATETIME_P.GetCommandParamValue("X_LOCAL_DATE"));

            //반려대상 선택.
            if (Set_Update_Return(vLOCAL_DATE) == false)
            {
                return;
            }

            HRMF0339_RETURN vHRMF0339_RETURN = new HRMF0339_RETURN(isAppInterfaceAdv1.AppInterface
                                                                  , WORK_CORP_ID_0.EditValue
                                                                  , vLOCAL_DATE
                                                                  );
            dlgResultValue = vHRMF0339_RETURN.ShowDialog();
            if (dlgResultValue == DialogResult.OK)
            {
            }
            vHRMF0339_RETURN.Dispose();

            Search_DB();

            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
        }


        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (IGR_DAY_WORK_MODIFY_APPR.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_APPROVE_YN = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN");
            int vIDX_PERSON_ID = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("PERSON_ID");
            int vIDX_WORK_DATE = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("WORK_DATE");
            int vIDX_APPROVE_STATUS = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DAY_WORK_MODIFY_APPR.RowCount; i++)
            {
                if (iConv.ISNull(IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_APPROVE_YN), "N") == "Y")
                {
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_PERSON_ID", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_PERSON_ID));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_WORK_DATE", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_WORK_DATE));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_STATUS", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_APPROVE_STATUS));  
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_APPROVE_YN));
                    IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_FLAG", pApproved_Flag);
                    IDC_SET_UPDATE_APPROVE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
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

            IDA_DAY_WORK_MODIFY_APPR.Cancel();
            Search_DB();
        }



        private bool Set_Update_Return(DateTime pSys_Date)
        {
            if (IGR_DAY_WORK_MODIFY_APPR.RowCount < 1)
            {
                return false;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            IGR_DAY_WORK_MODIFY_APPR.LastConfirmChanges();
            IDA_DAY_WORK_MODIFY_APPR.OraSelectData.AcceptChanges();
            IDA_DAY_WORK_MODIFY_APPR.Refillable = true; 

            int vIDX_SELECT_YN = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN");
            int vIDX_WORK_DATE = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("WORK_DATE"); 
            int vIDX_PERSON_ID = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("PERSON_ID");
            int vIDX_APPROVE_STATUS = IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("APPROVE_STATUS");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DAY_WORK_MODIFY_APPR.RowCount; i++)
            {
                if (iConv.ISNull(IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_SELECT_YN), "N") == "Y")
                { 
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_CHECK_YN", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_SELECT_YN));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_WORK_DATE", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_WORK_DATE)); 
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_PERSON_ID", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_PERSON_ID));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_APPROVE_STATUS", IGR_DAY_WORK_MODIFY_APPR.GetCellValue(i, vIDX_APPROVE_STATUS));
                    IDC_UPDATE_RETURN_TEMP.SetCommandParamValue("P_SYS_DATE", pSys_Date);
                    IDC_UPDATE_RETURN_TEMP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_UPDATE_RETURN_TEMP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_UPDATE_RETURN_TEMP.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        Application.DoEvents();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return false;
                    }
                }
            }
            return true;
        }


        private void Send_Mail()
        {
            IDC_GetDate.ExecuteNonQuery();
            object vLOCAL_DATE = iDate.ISGetDate(IDC_GetDate.GetCommandParamValue("X_LOCAL_DATE")).ToShortDateString();

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "DUTY");
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", vLOCAL_DATE);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

        #region ----- Grid Event -----

        private void igrDAY_INTERFACE_CellDoubleClick(object pSender)
        {
            if (IGR_DAY_WORK_MODIFY_APPR.RowIndex < 0 && IGR_DAY_WORK_MODIFY_APPR.ColIndex == IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN"))
            {
                for (int r = 0; r < IGR_DAY_WORK_MODIFY_APPR.RowCount; r++)
                {
                    if (iConv.ISNull(IGR_DAY_WORK_MODIFY_APPR.GetCellValue(r, IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN")), "N") == "Y".ToString())
                    {
                        IGR_DAY_WORK_MODIFY_APPR.SetCellValue(r, IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN"), "N");
                    }
                    else
                    {
                        IGR_DAY_WORK_MODIFY_APPR.SetCellValue(r, IGR_DAY_WORK_MODIFY_APPR.GetColumnToIndex("SELECT_YN"), "Y");
                    }
                }
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation Name(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDAY_INTERFACE_PreDelete(ISPreDeleteEventArgs e)
        {
        }

        private void idaDAY_INTERFACE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(IGR_DAY_WORK_MODIFY_APPR.GetCellValue("PERSON_ID"), IGR_DAY_WORK_MODIFY_APPR.GetCellValue("WORK_DATE"));
        }

        private void idaDAY_INTERFACE_UpdateCompleted(object pSender)
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "WORK");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", WORK_CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

        #region ----- LookUp Event -----
        
        private void ildWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaFLOOR_0_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_MODIFY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", W_WORK_DATE_TO.EditValue);
        }

        private void ilaPERSON_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", W_WORK_DATE_TO.EditValue);
        }

        #endregion

    }
}