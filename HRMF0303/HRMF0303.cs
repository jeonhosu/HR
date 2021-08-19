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

namespace HRMF0303
{
    public partial class HRMF0303 : Office2007Form
    {
        ISFunction.ISDateTime ISDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0303(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iString.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)   //파견직관리
            {
                G_CORP_TYPE_0.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
                G_CORP_TYPE_1.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
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

            CORP_NAME_0.BringToFront();
            CORP_NAME_1.BringToFront();

            igbCORP_GROUP_0.BringToFront();
            igbCORP_GROUP_1.BringToFront();

            igbCORP_GROUP_0.Visible = false; //.Show();
            igbCORP_GROUP_1.Visible = false;

            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iString.ISNull(G_CORP_TYPE_0.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();
                igbCORP_GROUP_1.Visible = true;

                irb_ALL_0.RadioButtonValue = "A";
                irb_ALL_1.RadioButtonValue = "A";
                G_CORP_TYPE_0.EditValue = "A";
                G_CORP_TYPE_1.EditValue = "A";
            }
            else if (iString.ISNull(G_CORP_TYPE_0.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                //CORP_NAME_1.EditValue = CORP_NAME_0.EditValue;
                //CORP_ID_1.EditValue = CORP_ID_0.EditValue;
            }
           
        }

        private void isSEARCH_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == 1)
            {
                if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty && iString.ISNull(G_CORP_TYPE_0.EditValue) != "4")
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CORP_NAME_0.Focus();
                    return;
                }
                if (iString.ISNull(WORK_YYYYMM_0.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WORK_YYYYMM_0.Focus();
                    return;
                }

                isAppInterfaceAdv1.OnAppMessage("");

                idaWORKCALENDAR.OraSelectData.AcceptChanges();
                idaWORKCALENDAR.Refillable = true;
                igrWORK_CALENDAR.LastConfirmChanges();

                idaPERSON_INFO.Fill();
                igrPERSON_INFO.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == 2)
            {
                if (iString.ISNull(CORP_ID_1.EditValue) == string.Empty && iString.ISNull(G_CORP_TYPE_1.EditValue) != "4")
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CORP_NAME_1.Focus();
                    return;
                }
                if (iString.ISNull(WORK_YYYYMM_1.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WORK_YYYYMM_1.Focus();
                    return;
                }

                IDA_NOT_CREATE_CALENDAR.Fill();
                IGR_NOT_CREATE_CALENDAR.Focus();
            }
        }

        private void Person_Info()
        {
            PERSON_NAME.EditValue = igrPERSON_INFO.GetCellValue("NAME");
            JOB_CATEGORY_NAME.EditValue = igrPERSON_INFO.GetCellValue("JOB_CATEGORY_NAME");
            JOIN_DATE.EditValue = igrPERSON_INFO.GetCellValue("JOIN_DATE");
            RETIRE_DATE.EditValue = igrPERSON_INFO.GetCellValue("RETIRE_DATE");
            START_DATE.EditValue = igrPERSON_INFO.GetCellValue("START_DATE");
            END_DATE.EditValue = igrPERSON_INFO.GetCellValue("END_DATE");
        }

        private void ISCalendarCreated(string pForm_ID)
        {
            if (pForm_ID == "HRMF0303_CREATE")
            {
                //MessageBoxAdv.Show("생성완료", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string  ISCap_Check()
        {// 근무계획표 변경 권한 체크.
            string sCap_Level;
            idcCAP_LEVEL.SetCommandParamValue("W_MODULE_CODE", "20".ToString());
            idcCAP_LEVEL.ExecuteNonQuery();
            sCap_Level = idcCAP_LEVEL.GetCommandParamValue("O_CAP_LEVEL").ToString();
            return sCap_Level;
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
                return true;
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
            return(mCheck_Value);
        }

        private void Init_Work_Time(object pWork_Type, object pHoly_Type)
        { 
            object mOPEN_TIME;
            object mCLOSE_TIME;
            idcWORK_IO_TIME.SetCommandParamValue("W_WORK_TYPE", pWork_Type);
            idcWORK_IO_TIME.SetCommandParamValue("W_HOLY_TYPE", pHoly_Type);
            idcWORK_IO_TIME.ExecuteNonQuery();
            mOPEN_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME");
            mCLOSE_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME");
            igrWORK_CALENDAR.SetCellValue("OPEN_TIME", mOPEN_TIME);
            igrWORK_CALENDAR.SetCellValue("CLOSE_TIME", mCLOSE_TIME);
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    isSEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaWORKCALENDAR.IsFocused)
                    {
                        // 권한 체크
                        if (ISCap_Check() != "C".ToString())
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10009", "&&CAP:=Create Work Calendar(근무계획표 생성)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaPERSON_INFO.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaWORKCALENDAR.IsFocused)
                    {
                        idaWORKCALENDAR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaWORKCALENDAR.IsFocused)
                    {
                        idaWORKCALENDAR.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0303_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0303_Shown(object sender, EventArgs e)
        {
            // 조회년월 SETTING
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");

            WORK_YYYYMM_0.EditValue = ISDate.ISYearMonth(DateTime.Today);
            idcYYYYMM_TERM.SetCommandParamValue("W_YYYYMM", WORK_YYYYMM_0.EditValue);
            idcYYYYMM_TERM.ExecuteNonQuery();
            START_DATE_0.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_START_DATE");
            END_DATE_0.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_END_DATE");

            WORK_YYYYMM_1.EditValue = WORK_YYYYMM_0.EditValue;
            START_DATE_1.EditValue = START_DATE_0.EditValue;
            END_DATE_1.EditValue = END_DATE_0.EditValue;

            DefaultCorporation();

            idaPERSON_INFO.FillSchema();
        }

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE_0.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void irb_ALL_1_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE_1.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void BTN_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(CORP_ID_0.EditValue) == null && iString.ISNull(G_CORP_TYPE_0.EditValue) != "4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(WORK_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_YYYYMM_0.Focus();
                return;
            }

            // 권한 체크
            if (ISCap_Check() != "C".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10009", "&&CAP:=Create Work Calendar(근무계획표 생성)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            System.Windows.Forms.DialogResult vdlgResultValue;
            HRMF0303_UPLOAD vHRMF0303_UPLOAD = new HRMF0303_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , CORP_ID_0.EditValue, WORK_YYYYMM_0.EditValue);
            vdlgResultValue = vHRMF0303_UPLOAD.ShowDialog();
            vHRMF0303_UPLOAD.Dispose();
        }

        private void ibtCALENDAR_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty && iString.ISNull(G_CORP_TYPE_0.EditValue) != "4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iString.ISNull(WORK_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 권한 체크
            if (ISCap_Check() != "C".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10009", "&&CAP:=Create Work Calendar(근무계획표 생성)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            System.Windows.Forms.DialogResult vdlgResultValue;
            Form vHRMF0303_CREATE = new HRMF0303_CREATE(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                            CORP_ID_0.EditValue, WORK_YYYYMM_0.EditValue, G_CORP_TYPE_0.EditValue);
            vdlgResultValue = vHRMF0303_CREATE.ShowDialog();
            vHRMF0303_CREATE.Dispose();
        }

        private void igrWORK_CALENDAR_CellDoubleClick(object pSender)
        {
            if(igrWORK_CALENDAR.RowIndex < 0)
            {
                return;
            }

            object mHoly_Type = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            object mWork_Type = igrWORK_CALENDAR.GetCellValue("WORK_TYPE");
            Init_Work_Time(mWork_Type, mHoly_Type);
        }

        private void igrWORK_CALENDAR_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            object mHoly_Type = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            object mWork_Date = igrWORK_CALENDAR.GetCellValue("WORK_DATE");
            object mNew_Work_Date;

            if (e.ColIndex == igrWORK_CALENDAR.GetColumnToIndex("OPEN_TIME"))
            {// 출근일자
                mNew_Work_Date = e.NewValue;
                if (Check_Work_Date_time(mHoly_Type, "IN".ToString(), mWork_Date, mNew_Work_Date) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            if (e.ColIndex == igrWORK_CALENDAR.GetColumnToIndex("CLOSE_TIME"))
            {// 출근일자
                mNew_Work_Date = e.NewValue;
                if (Check_Work_Date_time(mHoly_Type, "OUT".ToString(), mWork_Date, mNew_Work_Date) == false)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            if(e.ColIndex == igrWORK_CALENDAR.GetColumnToIndex("C_HOLY_TYPE_NAME"))
            {
                object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
                if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1")) != String.Empty)
                {
                    vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
                }
                else if (iString.ISNull(e.NewValue) != String.Empty)
                {
                    vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
                }
                Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE);
            }
            else if(e.ColIndex == igrWORK_CALENDAR.GetColumnToIndex("C_HOLY_TYPE_NAME1"))
            {
                object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
                if (iString.ISNull(e.NewValue) != String.Empty)
                {
                    vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
                }
                else if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE")) != String.Empty)
                {
                    vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
                }
                Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE);
            }
        }
      
        #endregion

        #region ----- Adapter Event -----

        private void idaWORKCALENDAR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Name(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무 일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DUTY_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Name(근태 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["HOLY_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy Type(근무 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            string vHOLY_TYPE = iString.ISNull(e.Row["HOLY_TYPE"]);
            if (iString.ISNull(e.Row["C_HOLY_TYPE1"]) != String.Empty)
            {
                vHOLY_TYPE = iString.ISNull(e.Row["C_HOLY_TYPE1"]);
            }
            else if (iString.ISNull(e.Row["C_HOLY_TYPE"]) != String.Empty)
            {
                vHOLY_TYPE = iString.ISNull(e.Row["C_HOLY_TYPE"]);
            }
            if ((vHOLY_TYPE == "H".ToString() || vHOLY_TYPE == "C".ToString() ||
                    vHOLY_TYPE == "1".ToString() || vHOLY_TYPE == "0".ToString())) 
            {// 휴일
                if (e.Row["OPEN_TIME"] != DBNull.Value || e.Row["CLOSE_TIME"] != DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10040", "&&VALUE:=Plan Open Time/Plan Close Time(계획 출근/퇴근 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            else
            {
                if (e.Row["OPEN_TIME"] == DBNull.Value && e.Row["CLOSE_TIME"] == DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Plan Open Time/Plan Close Time(계획 출근/퇴근 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "IN".ToString(), e.Row["WORK_DATE"], e.Row["OPEN_TIME"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (Check_Work_Date_time(e.Row["HOLY_TYPE"], "OUT".ToString(), e.Row["WORK_DATE"], e.Row["CLOSE_TIME"]) == false)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10151"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaWORKCALENDAR_PreDelete(ISPreDeleteEventArgs e)
        {
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            e.Cancel = true;
            return;
        }
        #endregion

        #region ----- LookUP Event -----

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_WORK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_WORK_TYPE_SelectedRowData(object pSender)
        {
            object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
            }
            else if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
            }
            Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE);
        }

        private void ilaDUTY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaHOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaHOLY_TYPE_SelectedRowData(object pSender)
        {
            object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
            }
            else if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
            }
            Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE);
        }

        private void ilaC_HOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaC_HOLY_TYPE_SelectedRowData(object pSender)
        {
            object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            if(iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
            }
            else if(iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
            } 
            Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE); 
        }

        private void ilaC_HOLY_TYPE1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaC_HOLY_TYPE1_SelectedRowData(object pSender)
        {
            object vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("HOLY_TYPE");
            if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE1");
            }
            else if (iString.ISNull(igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE")) != String.Empty)
            {
                vHOLY_TYPE = igrWORK_CALENDAR.GetCellValue("C_HOLY_TYPE");
            }
            Init_Work_Time(igrWORK_CALENDAR.GetCellValue("WORK_TYPE"), vHOLY_TYPE);
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_DEPT_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ILA_WORK_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        #endregion

    }
}