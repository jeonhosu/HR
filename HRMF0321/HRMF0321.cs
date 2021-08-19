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

namespace HRMF0321
{
    public partial class HRMF0321 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0321(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
            G_CORP_GROUP.BringToFront(); 
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iConv.ISNull(G_CORP_TYPE.EditValue, "1") == "1")
            {
                G_CORP_GROUP.Visible = false; //.Show();
                V_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G_CORP_TYPE.EditValue = V_RB_OWNER.RadioCheckedString;
            }
            else
            {
                G_CORP_GROUP.Visible = true; //.Show();
                if (iConv.ISNull(G_CORP_TYPE.EditValue) == "ALL")
                {
                    V_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ALL.RadioCheckedString;
                }
                else
                {
                    V_RB_ETC.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ETC.RadioCheckedString;
                }
            } 
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

            igrDAY_INTERFACE.LastConfirmChanges();
            idaDAY_INTERFACE.OraSelectData.AcceptChanges();
            idaDAY_INTERFACE.Refillable = true;

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

            if (iConv.ISNull(pHoly_Type) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(IO_Flag) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(pWork_Date) == string.Empty)
            {
                return (mCheck_Value);
            }
            if (iConv.ISNull(pNew_Work_Date) == string.Empty)
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
                        idaDAY_INTERFACE.SetUpdateParamValue("P_CONNECT_LEVEL", "C");
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

        private void HRMF0321_Load(object sender, EventArgs e)
        {
            WORK_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]    

            idaDAY_INTERFACE.FillSchema();            
        }

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void ibtnSET_DAY_INTERFACE_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 출퇴근 집계
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iConv.ISNull(WORK_DATE_0.EditValue) == string.Empty)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }
            
            idcSET_DAY_INTERFACE.SetCommandParamValue("P_CONNECT_LEVEL", "C");
            idcSET_DAY_INTERFACE.ExecuteNonQuery();
            string mSTATUS = iConv.ISNull(idcSET_DAY_INTERFACE.GetCommandParamValue("O_STATUS"));
            string mMessage = iConv.ISNull(idcSET_DAY_INTERFACE.GetCommandParamValue("O_MESSAGE"));

            if (idcSET_DAY_INTERFACE.ExcuteError)
            {
                MessageBoxAdv.Show(idcSET_DAY_INTERFACE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // refill.
            Search_DB();
        }

        private void BTN_EXCEL_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            HRMF0321_UPLOAD vHRMF0321_UPLOAD = new HRMF0321_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            vHRMF0321_UPLOAD.ShowDialog();
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
            if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") ||
                e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME1") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME1"))
            {
                object mHoly_Type = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
                object mWork_Date = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
                object mModify_Time = e.NewValue;

                object mIO_Flag = "-";
                if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME1"))
                {
                    mIO_Flag = "IN";
                }
                else if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME1"))
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
            if (igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") == igrDAY_INTERFACE.ColIndex)
            {
                if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue("OPEN_TIME")) == string.Empty)
                {
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", igrDAY_INTERFACE.GetCellValue("WORK_TYPE_GROUP"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", igrDAY_INTERFACE.GetCellValue("HOLY_TYPE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_OPEN_TIME", null);
                    idcWORK_IO_TIME.ExecuteNonQuery();
                    //출근
                    igrDAY_INTERFACE.SetCellValue("OPEN_TIME", idcWORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME"));
                }
            }
            if (igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") == igrDAY_INTERFACE.ColIndex)
            {
                if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue("CLOSE_TIME")) == string.Empty)
                {
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", igrDAY_INTERFACE.GetCellValue("WORK_TYPE_GROUP"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", igrDAY_INTERFACE.GetCellValue("HOLY_TYPE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
                    idcWORK_IO_TIME.SetCommandParamValue("W_OPEN_TIME", igrDAY_INTERFACE.GetCellValue("OPEN_TIME"));
                    idcWORK_IO_TIME.ExecuteNonQuery();
                    //퇴근
                    igrDAY_INTERFACE.SetCellValue("CLOSE_TIME", idcWORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME"));
                }
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["CORP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
             
            if ((iConv.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
                && (iConv.ISNull(e.Row["OPEN_TIME"]) == iConv.ISNull(e.Row["IN_TIME"])
                && iConv.ISNull(e.Row["OPEN_TIME1"]) == iConv.ISNull(e.Row["IN_TIME1"]))
                && (iConv.ISNull(e.Row["CLOSE_TIME"]) == iConv.ISNull(e.Row["OUT_TIME"])
                && iConv.ISNull(e.Row["CLOSE_TIME1"]) == iConv.ISNull(e.Row["OUT_TIME1"])))
            {
            }
            else if ((iConv.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
            && (iConv.ISNull(e.Row["OPEN_TIME"]) == string.Empty)
            && (iConv.ISNull(e.Row["OPEN_TIME1"]) == string.Empty))
            {
            }
            else
            {
                if (iConv.ISNull(e.Row["OPEN_TIME"]) != iConv.ISNull(e.Row["IN_TIME"])
                    || iConv.ISNull(e.Row["OPEN_TIME1"]) != iConv.ISNull(e.Row["IN_TIME1"])
                    || iConv.ISNull(e.Row["CLOSE_TIME"]) != iConv.ISNull(e.Row["OUT_TIME"])
                    || iConv.ISNull(e.Row["CLOSE_TIME1"]) != iConv.ISNull(e.Row["OUT_TIME1"]))
                {// 시간 변경이 있을 경우.
                    if (iConv.ISNull(e.Row["MODIFY_ID"]) == string.Empty)
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

            if (iConv.ISNull(e.Row["LEAVE_ID"]) != string.Empty && iConv.ISNull(e.Row["LEAVE_TIME_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10255"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(e.Row["LEAVE_ID"]) == string.Empty && iConv.ISNull(e.Row["LEAVE_TIME_CODE"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10254"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            
            object mIO_Flag = "-".ToString();
            if (iConv.ISNull(e.Row["OPEN_TIME"]) != iConv.ISNull(e.Row["IN_TIME"])
                || iConv.ISNull(e.Row["OPEN_TIME1"]) != iConv.ISNull(e.Row["IN_TIME1"]))
            {
                mIO_Flag = "IN";
                if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["OPEN_TIME"]) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["OPEN_TIME1"]) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

            if (iConv.ISNull(e.Row["CLOSE_TIME"]) != iConv.ISNull(e.Row["OUT_TIME"])
                || iConv.ISNull(e.Row["CLOSE_TIME1"]) != iConv.ISNull(e.Row["OUT_TIME1"]))
            {
                mIO_Flag = "OUT";
                if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["CLOSE_TIME"]) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (Check_Work_Date_time(e.Row["HOLY_TYPE"], mIO_Flag, e.Row["WORK_DATE"], e.Row["CLOSE_TIME1"]) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
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

        private void ILA_W_OPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
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

        private void ilaYES_NO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YES_NO");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}