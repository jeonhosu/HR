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

namespace HRMF0309
{
    public partial class HRMF0309 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0309(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

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
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }
            if (String.IsNullOrEmpty(INOUT_FLAG.EditValue.ToString()))
            {// 출퇴구분 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty In/Out Type(출퇴구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (INOUT_FLAG.EditValue.ToString() == "1".ToString())
            {// 출근
                // 후일 퇴근.
                igrDAY_INTERFACE.GridAdvExColElement[14].Insertable = 0;
                igrDAY_INTERFACE.GridAdvExColElement[14].Updatable = 0;
                // 당직.
                igrDAY_INTERFACE.GridAdvExColElement[15].Insertable = 0;
                igrDAY_INTERFACE.GridAdvExColElement[15].Updatable = 0;
                //철야
                igrDAY_INTERFACE.GridAdvExColElement[16].Insertable = 0;
                igrDAY_INTERFACE.GridAdvExColElement[16].Updatable = 0;

                //외출시간.
                igrDAY_INTERFACE.GridAdvExColElement[18].Insertable = 0;
                igrDAY_INTERFACE.GridAdvExColElement[18].Updatable = 0;

                //외출사유.
                igrDAY_INTERFACE.GridAdvExColElement[20].Insertable = 0;
                igrDAY_INTERFACE.GridAdvExColElement[20].Updatable = 0;

            }
            else
            {// 퇴근
                // 후일 퇴근.
                igrDAY_INTERFACE.GridAdvExColElement[14].Insertable = 1;
                igrDAY_INTERFACE.GridAdvExColElement[14].Updatable = 1;
                // 당직.
                igrDAY_INTERFACE.GridAdvExColElement[15].Insertable = 1;
                igrDAY_INTERFACE.GridAdvExColElement[15].Updatable = 1;
                //철야
                igrDAY_INTERFACE.GridAdvExColElement[16].Insertable = 1;
                igrDAY_INTERFACE.GridAdvExColElement[16].Updatable = 1;

                //외출시간.
                igrDAY_INTERFACE.GridAdvExColElement[18].Insertable = 1;
                igrDAY_INTERFACE.GridAdvExColElement[18].Updatable = 1;

                //외출사유.
                igrDAY_INTERFACE.GridAdvExColElement[20].Insertable = 1;
                igrDAY_INTERFACE.GridAdvExColElement[20].Updatable = 1;
            }
            
            idaDAY_INTERFACE.Fill();
            igrDAY_INTERFACE.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = WORK_DATE_0.EditValue;

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

        private bool Check_Work_Date_time(object pHoly_Type, object IO_Flag, object pWork_Date, object pNew_Work_Date)
        {
            bool mCheck_Value = false;

            if (iString.ISNull(pHoly_Type) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iString.ISNull(IO_Flag) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iString.ISNull(pWork_Date) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iString.ISNull(pNew_Work_Date) == string.Empty)
            {
                return (true);
            }

            if ((pHoly_Type.ToString() == "0".ToString() || pHoly_Type.ToString() == "1".ToString() || pHoly_Type.ToString() == "2".ToString()
                || pHoly_Type.ToString() == "D".ToString() || pHoly_Type.ToString() == "S".ToString())
                && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date == Convert.ToDateTime(pNew_Work_Date).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() || pHoly_Type.ToString() == "N".ToString())
                && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 야간, 무휴, 유휴, DAY, NIGHT --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "0".ToString() || pHoly_Type.ToString() == "1".ToString() || pHoly_Type.ToString() == "2".ToString()
         || pHoly_Type.ToString() == "D".ToString() || pHoly_Type.ToString() == "S".ToString())
              && IO_Flag.ToString() == "OUT".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() || pHoly_Type.ToString() == "N".ToString())
           && IO_Flag.ToString() == "OUT".ToString())
            {// 주간, 야간, 무휴, 유휴, DAY, NIGHT --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                    && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            return (mCheck_Value);
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
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.SetUpdateParamValue("P_CONNECT_LEVEL", "A");
                        idaDAY_INTERFACE.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----
        private void HRMF0309_Load(object sender, EventArgs e)
        {
            idaDAY_INTERFACE.FillSchema();
            WORK_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            irbIN.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        private void ibtnSET_DAY_INTERFACE_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 출퇴근 집계
            string mMessage;

            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }

            idcSET_DAY_INTERFACE.ExecuteNonQuery();
            mMessage = idcSET_DAY_INTERFACE.GetCommandParamValue("O_MESSAGE").ToString();
            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // refill.
            Search_DB();
        }

        private void btnAPPR_REQUEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaDAY_INTERFACE.Update();

            int mRowCount = igrDAY_INTERFACE.RowCount;
            for (int R = 0; R < mRowCount; R++)
            {
                if (iString.ISNull(igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"))) == "N".ToString())
                {// 승인미요청 건에 대해서 승인 처리.
                    idcAPPROVAL_REQUEST.SetCommandParamValue("W_PERSON_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("PERSON_ID")));
                    idcAPPROVAL_REQUEST.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("WORK_DATE")));
                    idcAPPROVAL_REQUEST.SetCommandParamValue("W_CORP_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("CORP_ID")));
                    idcAPPROVAL_REQUEST.ExecuteNonQuery();

                    object mValue;
                    mValue = idcAPPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                    igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"), mValue);
                    mValue = idcAPPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                    igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                }
            }

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "WORK");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", WORK_DATE_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
            
            idaDAY_INTERFACE.OraSelectData.AcceptChanges();
            idaDAY_INTERFACE.Refillable = true;
        }

        private void irbINOUT_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv isINOUT = sender as ISRadioButtonAdv;
            INOUT_FLAG.EditValue = isINOUT.RadioCheckedString;

            // refill.
            Search_DB();
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }

        private void igrDAY_INTERFACE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("MODIFY_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("MODIFY_TIME1"))
            {
                object mHoly_Type = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
                object mIO_Flag = igrDAY_INTERFACE.GetCellValue("IO_FLAG");
                object mWork_Date = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
                object mModify_Time = e.NewValue;

                if (iString.ISNull(mIO_Flag) == "1".ToString())
                {
                    mIO_Flag = "IN";
                }
                else if (iString.ISNull(mIO_Flag) == "2".ToString())
                {
                    mIO_Flag = "OUT";
                }
                if (Check_Work_Date_time(mHoly_Type, mIO_Flag, mWork_Date, mModify_Time) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }                
            }
        }

        private void igrDAY_INTERFACE_CellDoubleClick(object pSender)
        {
            if (igrDAY_INTERFACE.GetColumnToIndex("MODIFY_TIME") == igrDAY_INTERFACE.ColIndex)
            {
                if (iString.ISNull(igrDAY_INTERFACE.GetCellValue("MODIFY_TIME")) == string.Empty)
                {
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", igrDAY_INTERFACE.GetCellValue("WORK_TYPE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", igrDAY_INTERFACE.GetCellValue("HOLY_TYPE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
                    idcWORK_IO_TIME.ExecuteNonQuery();
                    if (iString.ISNull(igrDAY_INTERFACE.GetCellValue("IO_FLAG")) == "1".ToString())
                    {//출근
                        igrDAY_INTERFACE.SetCellValue("MODIFY_TIME", idcWORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME"));
                    }
                    else if (iString.ISNull(igrDAY_INTERFACE.GetCellValue("IO_FLAG")) == "2".ToString())
                    {//퇴근
                        igrDAY_INTERFACE.SetCellValue("MODIFY_TIME", idcWORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME"));
                    }
                }
            }
        }

        #endregion  

        #region ----- Adapter Event -----
        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iString.ISNull(e.Row["IO_FLAG"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty In/Out Flag(출퇴구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if ((iString.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
                && (iString.ISNull(e.Row["MODIFY_TIME"]) == iString.ISNull(e.Row["IO_TIME"])
                && iString.ISNull(e.Row["MODIFY_TIME1"]) == iString.ISNull(e.Row["IO_TIME1"])))
            {
            }
            else if ((iString.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
            && (iString.ISNull(e.Row["MODIFY_TIME"]) == string.Empty)
            && (iString.ISNull(e.Row["MODIFY_TIME1"]) == string.Empty))
            {
            }
            else
            {
                if (iString.ISNull(e.Row["MODIFY_TIME"]) != iString.ISNull(e.Row["IO_TIME"])
                    || iString.ISNull(e.Row["MODIFY_TIME1"]) != iString.ISNull(e.Row["IO_TIME1"]))
                {// 시간 변경이 있을 경우.
                    if (iString.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
                    {// 수정 사유 체크
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Modify Reason(수정사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }
                }
                //if ((iString.ISNull(e.Row["MODIFY_ID"]) != string.Empty)
                //    && (iString.ISNull(e.Row["MODIFY_TIME"]) == string.Empty
                //    && iString.ISNull(e.Row["MODIFY_TIME1"]) == string.Empty))
                //{// 시간 변경이 있을 경우.                
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Modify Reason(수정사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    e.Cancel = true;
                //    return;
                //}
            }

            object mIO_Flag = "-".ToString();
            if (iString.ISNull(e.Row["IO_FLAG"]) == "1".ToString())
            {
                mIO_Flag = "IN";
            }
            else if (iString.ISNull(e.Row["IO_FLAG"]) == "2".ToString())
            {
                mIO_Flag = "OUT";
            }
            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["MODIFY_TIME"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["MODIFY_TIME1"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDAY_INTERFACE_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        private void idaDAY_INTERFACE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_INTERFACE.GetCellValue("PERSON_ID"), igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
        }

        #endregion

        #region ----- LookUp Event -----
        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }
        
        private void ildWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", WORK_DATE_0.EditValue);
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_MODIFY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaLEAVE_OUT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "LEAVE_OUT");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaLEAVE_OUT_TIME_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "LEAVE_OUT_TIME");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion
        
    }
}