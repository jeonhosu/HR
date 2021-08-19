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

namespace HRMF0388
{
    public partial class HRMF0388 : Office2007Form
    {        
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISDateTime iSDate = new ISCommonUtil.ISFunction.ISDateTime();
        private ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0388(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE_0.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }

        #endregion;

        #region ----- Private Methods ----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildWORK_CORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildWORK_CORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            

            WORK_CORP_NAME_0.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            igbCORP_GROUP_0.Visible = false;

            if (iConv.ISNull(CORP_TYPE_0.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();


                irb_ALL_0.RadioButtonValue = "A";
                CORP_TYPE_0.EditValue = "A";

            }
            else  
            {
                WORK_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                WORK_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            }


            
        }

        private void DefaultFloor()
        {
            //작업장
            IDC_DEFAULT_FLOOR.ExecuteNonQuery();
            FLOOR_NAME_0.EditValue = IDC_DEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            FLOOR_ID_0.EditValue = IDC_DEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");
        }

        private void Search_DB()
        {
            if (WORK_CORP_ID_0.EditValue == null && CORP_TYPE_0.EditValue.ToString() !="4")
            {// 업체. - 업체정보는 필수입니다. 선택하세요.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자 - 시작일자는 필수입니다
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }

            // 후일 퇴근.
            int vIndex_NEXT_DAY_YN = igrDAY_INTERFACE.GetColumnToIndex("NEXT_DAY_YN");
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_NEXT_DAY_YN].Insertable = 1;
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_NEXT_DAY_YN].Updatable = 1;

            // 당직.
            int vIndex_DANGJIK_YN = igrDAY_INTERFACE.GetColumnToIndex("DANGJIK_YN");
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_DANGJIK_YN].Insertable = 1;
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_DANGJIK_YN].Updatable = 1;

            //철야
            int vIndex_ALL_NIGHT_YN = igrDAY_INTERFACE.GetColumnToIndex("ALL_NIGHT_YN");
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_ALL_NIGHT_YN].Insertable = 1;
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_ALL_NIGHT_YN].Updatable = 1;

            //외출사유.
            int vIndex_LEAVE_NAME = igrDAY_INTERFACE.GetColumnToIndex("LEAVE_NAME");
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_LEAVE_NAME].Insertable = 1;
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_LEAVE_NAME].Updatable = 1;

            //외출시간.
            int vIndex_LEAVE_TIME = igrDAY_INTERFACE.GetColumnToIndex("LEAVE_TIME");
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_LEAVE_TIME].Insertable = 1;
            igrDAY_INTERFACE.GridAdvExColElement[vIndex_LEAVE_TIME].Updatable = 1;

            IDA_DAY_INTERFACE.Cancel();
            IDA_DAY_INTERFACE.OraSelectData.AcceptChanges();
            IDA_DAY_INTERFACE.Refillable = true;

            IDA_DAY_INTERFACE.Fill();
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

            if ((pHoly_Type.ToString() == "0".ToString() 
              || pHoly_Type.ToString() == "1".ToString() 
              || pHoly_Type.ToString() == "2".ToString()
              || pHoly_Type.ToString() == "D".ToString() 
              || pHoly_Type.ToString() == "S".ToString())
              && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date == Convert.ToDateTime(pNew_Work_Date).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() 
                   || pHoly_Type.ToString() == "N".ToString())
                   && IO_Flag.ToString() == "IN".ToString())
            {// 주간, 야간, 무휴, 유휴, DAY, NIGHT --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                 && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "0".ToString() 
                   || pHoly_Type.ToString() == "1".ToString() 
                   || pHoly_Type.ToString() == "2".ToString()
                   || pHoly_Type.ToString() == "D".ToString() 
                   || pHoly_Type.ToString() == "S".ToString())
                   && IO_Flag.ToString() == "OUT".ToString())
            {// 주간, 무휴, 유휴, DAY, SWING --> 같은 날짜.
                if (Convert.ToDateTime(pWork_Date).Date <= Convert.ToDateTime(pNew_Work_Date).Date
                 && Convert.ToDateTime(pNew_Work_Date).Date <= Convert.ToDateTime(pWork_Date).AddDays(1).Date)
                {
                    mCheck_Value = true;
                }
            }
            else if ((pHoly_Type.ToString() == "3".ToString() 
                   || pHoly_Type.ToString() == "N".ToString())
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
                    if (IDA_DAY_INTERFACE.IsFocused)
                    {
                        IDA_DAY_INTERFACE.SetUpdateParamValue("P_CONNECT_LEVEL", "A");
                        IDA_DAY_INTERFACE.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DAY_INTERFACE.IsFocused)
                    {
                        IDA_DAY_INTERFACE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DAY_INTERFACE.IsFocused)
                    {
                        IDA_DAY_INTERFACE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0388_Load(object sender, EventArgs e)
        {
            IDA_DAY_INTERFACE.FillSchema();

            WORK_DATE_0.EditValue = DateTime.Today;
            //WORK_DATE_0.EditValue = new System.DateTime(2011, 9, 2);

            irbAPPR_ERROR.CheckedState = ISUtil.Enum.CheckedState.Checked;            
            STATUS_TYPE.EditValue = irbAPPR_ERROR.RadioCheckedString;
            DefaultCorporation();
          //  DefaultFloor();
        }
         
        private void BTN_APPR_REQUEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                IDA_DAY_INTERFACE.Update();

                int mRowCount = igrDAY_INTERFACE.RowCount;
                for (int R = 0; R < mRowCount; R++)
                {
                    string vCHECK_FLAG = igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("SELECT_YN")).ToString();
                    if (vCHECK_FLAG == "Y")
                    {
                        if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"))) == "N".ToString())
                        {// 승인미요청 건에 대해서 승인요청 처리.
                            IDC_APPROVAL_REQUEST.SetCommandParamValue("W_PERSON_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("PERSON_ID")));
                            IDC_APPROVAL_REQUEST.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("WORK_DATE")));
                            IDC_APPROVAL_REQUEST.SetCommandParamValue("W_CORP_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("WORK_CORP_ID")));
                            IDC_APPROVAL_REQUEST.SetCommandParamValue("P_REQUEST_STATUS", "OK");
                            IDC_APPROVAL_REQUEST.ExecuteNonQuery();
                            string vSTATUS = iConv.ISNull(IDC_APPROVAL_REQUEST.GetCommandParamValue("O_STATUS"));
                            string vMESSAGE = iConv.ISNull(IDC_APPROVAL_REQUEST.GetCommandParamValue("O_MESSAGE"));
                            object mValue;
                            mValue = IDC_APPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS");
                            igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"), mValue);
                            mValue = IDC_APPROVAL_REQUEST.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                            igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                            if (vSTATUS == "F")
                            {
                                if(vMESSAGE != string.Empty)
                                {
                                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                return;
                            } 
                        }
                    }                    
                }

                // EMAIL 발송.
                idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
                idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "WORK");
                idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", WORK_CORP_ID_0.EditValue);
                idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", WORK_DATE_0.EditValue);
                idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
                idcEMAIL_SEND.ExecuteNonQuery();

                SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

                igrDAY_INTERFACE.LastConfirmChanges();
                IDA_DAY_INTERFACE.OraSelectData.AcceptChanges();
                IDA_DAY_INTERFACE.Refillable = true;

                //승인요청을 하셨습니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SKEAPP_10238"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBoxAdv.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BTN_APPR_R_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                //IDA_DAY_INTERFACE.Update();
                int mRowCount = igrDAY_INTERFACE.RowCount;
                for (int R = 0; R < mRowCount; R++)
                {
                    string vCHECK_FLAG = igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("SELECT_YN")).ToString();
                    if (vCHECK_FLAG == "Y")
                    {
                        if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"))) == "A".ToString())
                        {// 승인미요청 건에 대해서 승인요청 처리.
                            IDC_REQUEST_CANCEL.SetCommandParamValue("W_PERSON_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("PERSON_ID")));
                            IDC_REQUEST_CANCEL.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("WORK_DATE")));
                            IDC_REQUEST_CANCEL.SetCommandParamValue("W_CORP_ID", igrDAY_INTERFACE.GetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("WORK_CORP_ID")));
                            IDC_REQUEST_CANCEL.SetCommandParamValue("P_REQUEST_STATUS", "CANCEL");
                            IDC_REQUEST_CANCEL.ExecuteNonQuery();
                            string vSTATUS = iConv.ISNull(IDC_REQUEST_CANCEL.GetCommandParamValue("O_STATUS"));
                            string vMESSAGE = iConv.ISNull(IDC_REQUEST_CANCEL.GetCommandParamValue("O_MESSAGE"));
                            object mValue;
                            mValue = IDC_REQUEST_CANCEL.GetCommandParamValue("O_APPROVE_STATUS");
                            igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS"), mValue);
                            mValue = IDC_REQUEST_CANCEL.GetCommandParamValue("O_APPROVE_STATUS_NAME");
                            igrDAY_INTERFACE.SetCellValue(R, igrDAY_INTERFACE.GetColumnToIndex("APPROVE_STATUS_NAME"), mValue);
                            if (vSTATUS == "F")
                            {
                                if (vMESSAGE != string.Empty)
                                {
                                    MessageBoxAdv.Show(vMESSAGE, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                return;
                            }
                        }
                    }
                }
                igrDAY_INTERFACE.LastConfirmChanges();
                IDA_DAY_INTERFACE.OraSelectData.AcceptChanges();
                IDA_DAY_INTERFACE.Refillable = true;
                //승인요청취소를 하셨습니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("NFKEAPP_10231"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBoxAdv.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
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
            //if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME1"))
            //{
            //    object mHoly_Type = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
            //    object mIO_Flag = null;
            //    object mWork_Date = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
            //    object mModify_Time = e.NewValue;

            //    mIO_Flag = "IN";
            //    if (Check_Work_Date_time(mHoly_Type, mIO_Flag, mWork_Date, mModify_Time) == false)
            //    {
            //        //근무일자와 계획일자가 다릅니다. 확인하세요
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}
            //else if (e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") || e.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME1"))
            //{
            //    object mHoly_Type = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
            //    object mIO_Flag = null;
            //    object mWork_Date = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
            //    object mModify_Time = e.NewValue;

            //    mIO_Flag = "OUT";
            //    if (Check_Work_Date_time(mHoly_Type, mIO_Flag, mWork_Date, mModify_Time) == false)
            //    {
            //        //근무일자와 계획일자가 다릅니다. 확인하세요
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}
        }

        private void igrDAY_INTERFACE_CellDoubleClick(object pSender)
        {

            if (igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") == igrDAY_INTERFACE.ColIndex)
            {
                if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue("OPEN_TIME")) == string.Empty)
                {
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", igrDAY_INTERFACE.GetCellValue("WORK_TYPE_GROUP"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", igrDAY_INTERFACE.GetCellValue("HOLY_TYPE"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_OPEN_TIME", null);
                    IDC_WORK_IO_TIME.ExecuteNonQuery();
                    //출근
                    igrDAY_INTERFACE.SetCellValue("OPEN_TIME", IDC_WORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME"));
                }
            }
            if (igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") == igrDAY_INTERFACE.ColIndex)
            {
                if (iConv.ISNull(igrDAY_INTERFACE.GetCellValue("CLOSE_TIME")) == string.Empty)
                {
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", igrDAY_INTERFACE.GetCellValue("WORK_TYPE_GROUP"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", igrDAY_INTERFACE.GetCellValue("HOLY_TYPE"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
                    IDC_WORK_IO_TIME.SetCommandParamValue("W_OPEN_TIME", igrDAY_INTERFACE.GetCellValue("OPEN_TIME"));
                    IDC_WORK_IO_TIME.ExecuteNonQuery();
                    //퇴근
                    igrDAY_INTERFACE.SetCellValue("CLOSE_TIME", IDC_WORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME"));
                }
            }
            //object vWORK_GROUP = null;
            //object vHOLY_TYPE = null;
            //object vWORK_DATE = null;
            //object vDATE_TIME = null;

            //if (igrDAY_INTERFACE.GetColumnToIndex("OPEN_TIME") == igrDAY_INTERFACE.ColIndex)
            //{
            //    if (iString.ISNull(igrDAY_INTERFACE.GetCellValue("OPEN_TIME")) == string.Empty)
            //    {
            //        vWORK_GROUP = igrDAY_INTERFACE.GetCellValue("WORK_GROUP");
            //        vHOLY_TYPE = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
            //        vWORK_DATE = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
            //        idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", vWORK_GROUP);
            //        idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", vHOLY_TYPE);
            //        idcWORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", vWORK_DATE);
            //        idcWORK_IO_TIME.ExecuteNonQuery();

            //        vDATE_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME");
            //        igrDAY_INTERFACE.SetCellValue("OPEN_TIME", vDATE_TIME);
            //    }
            //}
            //else if (igrDAY_INTERFACE.GetColumnToIndex("CLOSE_TIME") == igrDAY_INTERFACE.ColIndex)
            //{
            //    if (iString.ISNull(igrDAY_INTERFACE.GetCellValue("CLOSE_TIME")) == string.Empty)
            //    {
            //        vWORK_GROUP = igrDAY_INTERFACE.GetCellValue("WORK_GROUP");
            //        vHOLY_TYPE = igrDAY_INTERFACE.GetCellValue("HOLY_TYPE");
            //        vWORK_DATE = igrDAY_INTERFACE.GetCellValue("WORK_DATE");
            //        idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", vWORK_GROUP);
            //        idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", vHOLY_TYPE);
            //        idcWORK_IO_TIME.SetCommandParamValue("W_WORK_DATE", vWORK_DATE);
            //        idcWORK_IO_TIME.ExecuteNonQuery();

            //        vDATE_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME");
            //        igrDAY_INTERFACE.SetCellValue("CLOSE_TIME", vDATE_TIME);
            //    }
            //}
        }



        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WORK_CORP_ID"]) == string.Empty && CORP_TYPE_0.EditValue.ToString() != "4")
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["MODIFY_ID"]) == string.Empty )
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:= Modify Description(수정사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
               
            //저장전 검증.
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("W_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("W_WORK_DATE", e.Row["WORK_DATE"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("W_CORP_ID", e.Row["CORP_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("W_DUTY_ID", e.Row["DUTY_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_OPEN_TIME", e.Row["OPEN_TIME"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_OPEN_TIME1", e.Row["OPEN_TIME1"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_CLOSE_TIME", e.Row["CLOSE_TIME"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_CLOSE_TIME1", e.Row["CLOSE_TIME1"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_MODIFY_ID", e.Row["MODIFY_ID"]); 
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_LEAVE_ID", e.Row["LEAVE_ID"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_LEAVE_TIME_CODE", e.Row["LEAVE_TIME_CODE"]); 
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_NEXT_DAY_YN", e.Row["NEXT_DAY_YN"]);
            IDC_VALIDATE_SAVE_P.SetCommandParamValue("P_DESCRIPTION", e.Row["DESCRIPTION"]);          
            IDC_VALIDATE_SAVE_P.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_VALIDATE_SAVE_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_VALIDATE_SAVE_P.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }    
        }

        private void idaDAY_INTERFACE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_INTERFACE.GetCellValue("PERSON_ID"), igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
        }

        #endregion

        #region ----- LOOKUP Event -----
        
        private void ildWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

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

        private void ilaFLOOR_0_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }


        #endregion

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_0.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void ilaPERSON_0_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", WORK_DATE_0.EditValue);
        }

        private void isCheckBoxAdv1_SystemColorsChanged(object sender, EventArgs e)
        {

        }

        private void isCheckBoxAdv1_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = igrDAY_INTERFACE.GetColumnToIndex("SELECT_YN");
            for (int r = 0; r < igrDAY_INTERFACE.RowCount; r++)
            {
                igrDAY_INTERFACE.SetCellValue(r, vIDX_SELECT_FLAG, SELECT_YN.CheckBoxString);
            }
        }

        private void irbAPPR_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            STATUS_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

     
    }
}