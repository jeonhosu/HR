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

namespace HRMF0380
{
    public partial class HRMF0380 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private bool misDelete = false;

        private bool mIsSwitch = false;

        private string mCAPACITY = string.Empty;


        //그리드 col 제어위해 그리드 col index 값 정의 
        private int mIDX_BF_START_TIME = 8;   //근무전 연장 시시간 
        private int mIDX_BF_END_TIME = 9;   //근무전 연장 시시간 
        private int mIDX_AF_START_DATE = 10;   //근무전 연장 시시간 
        private int mIDX_AF_START_TIME = 11;   //근무전 연장 시시간 

        #endregion;

        #region ----- Constructor -----

        public HRMF0380(Form pMainForm, ISAppInterface pAppInterface)
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

        private void GetCapacity()
        {
            try
            {
                idcGET_CAPACITY.ExecuteNonQuery();
                object oCAPACITY = idcGET_CAPACITY.GetCommandParamValue("O_CAPACITY_C");

                mCAPACITY = ConvertString(oCAPACITY);
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {
                //업체정보는 필수입니다. 선택하세요.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (STD_DATE_0.EditValue == null)
            {
                //조회년월을 선택하지 않았습니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE_0.Focus();
                return;
            }

            if (WORK_DATE.EditValue != null)
            {
                if (WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Saturday
                 || WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Sunday)
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
                }
                else
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 1; //수정 가능, 조기출근 시작시간
                }
            }


            ////주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화

            ////평일에는 근무후 시작일, 시작시를 수정할 일이 없어 수정 불가능하도록 설정

            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0;  //수정 불가능, 근무후 일자
            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능, 근무후 시간

            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
            //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;

            idaOT_HEADER.Fill();

            //-----------------------------------------------------------------------------
            //[2012-01-25]추가
            int vCountRow = igrOT_LINE.RowCount;

            if (vCountRow > 0)
            {
                int vIndexColumn_HOLY_TYPE_1 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_1");
                int vIndexColumn_HOLY_TYPE_2 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_2");

                object vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue(0, vIndexColumn_HOLY_TYPE_1);
                string vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                object vObject_HOLY_TYPE_2 = igrOT_LINE.GetCellValue(0, vIndexColumn_HOLY_TYPE_2);
                string vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

                //0:무급유일[토], 1:휴일[일]
                //주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화
                if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                {
                    //근무전연장 제어
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
                    //근무후 일자
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 1; //수정 가능
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 1;
                    //근무후 시간
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 1; //수정 가능
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 1;
                }
                else
                {
                    //근무전연장 제어
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 1; //수정 가능, 조기출근 시작시간

                    //평일에는 근무후 시작일, 시작시를 수정할 일이 없어 수정 불가능하도록 설정
                    //근무후 일자
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0; //수정 불가능
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
                    //근무후 시간
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;
                }
            }
            //-----------------------------------------------------------------------------
            igrOT_LINE.ResetDraw = true;
            igrOT_LINE.Refresh();
            igrOT_LINE.CurrentCellMoveTo(0, 0);
            igrOT_LINE.Focus();

            //REQ_NUM_0.Focus();
        }

        private bool isOT_Header_Check()
        {
            if (DUTY_MANAGER_ID_0.EditValue == null)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Floor Name(작업장)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return  true;
        }

        private void isOT_Header()
        {
            CORP_ID.EditValue = CORP_ID_0.EditValue;
            REQ_DATE.EditValue = iDate.ISGetDate();
            REQ_PERSON_ID.EditValue = isAppInterfaceAdv1.PERSON_ID;
            DUTY_MANAGER_ID.EditValue = DUTY_MANAGER_ID_0.EditValue;
            DUTY_MANAGER_NAME.EditValue = DUTY_MANAGER_NAME_0.EditValue;
            OT_HEADER_ID.EditValue = -1;

            idcDV_REQ_TYPE.SetCommandParamValue("W_GROUP_CODE", "REQ_TYPE");
            idcDV_REQ_TYPE.ExecuteNonQuery();
            REQ_TYPE.EditValue = idcDV_REQ_TYPE.GetCommandParamValue("O_CODE");
            REQ_TYPE_NAME.EditValue = idcDV_REQ_TYPE.GetCommandParamValue("O_CODE_NAME");
            
            WORK_DATE.EditValue = STD_DATE_0.EditValue;
            
            REQ_TYPE_NAME.Focus();

            if (WORK_DATE.EditValue != null)
            {
                if (WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Saturday
                 || WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Sunday)
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                }
                else
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 1; //수정 가능, 조기출근 시작시간
                }
            }
        }

        private bool isOT_Line_Check()
        {
            //bool vIsRequest = RequestLimit();
            //if (vIsRequest == false)
            //{
            //    return false;
            //}

            if (OT_HEADER_ID.EditValue == null)
            {
                //헤더정보를 찾을수 없습니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10239"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            //if (iString.ISNull(APPROVE_STATUS.EditValue, "N") != "A".ToString() )
            //{
            //    //해당 자료는 이미 승인되었습니다. &&VALUE 할 수 없습니다.
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10042", "&&VALUE:=Addition Request(추가신청)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return false;
            //}
            //else 
            if (iString.ISNull(APPROVE_STATUS.EditValue, "N") != "N".ToString() && iString.ISNull(REJECT_YN.EditValue) == "N".ToString())
            {
                //해당 자료는 이미 승인되었습니다. &&VALUE 할 수 없습니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10506"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }        

            return true;
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
            return (mCheck_Value);
        }

        #endregion;

        #region ----- Get Date Method -----

        private bool Set_Request (string pStatus)
        {
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            idcSET_UPDATE_REQUEST.SetCommandParamValue("P_REQUEST_STATUS", pStatus);
            idcSET_UPDATE_REQUEST.ExecuteNonQuery();
            vSTATUS = iString.ISNull(idcSET_UPDATE_REQUEST.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(idcSET_UPDATE_REQUEST.GetCommandParamValue("O_MESSAGE"));
            if (idcSET_UPDATE_REQUEST.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }
            return true;
        }

        #endregion;


        #region ----- Get Date Method -----

        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                idcGetDate.ExecuteNonQuery();
                object vObject = idcGetDate.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        #endregion;

        #region ----- Convert decimal  Method ----

        private decimal ConvertNumber(object pObject)
        {
            bool vIsConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        vConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertDecimal;
        }

        #endregion;

        #region ----- Convert String Method ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- Convert DateTime Methods ----

        private System.DateTime ConvertDateTime(object pObject)
        {
            System.DateTime vDateTime = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        vDateTime = (System.DateTime)pObject;
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vDateTime;
        }

        #endregion;

        #region ----- Convert Date Method ----

        private string ConvertDate(object pObject)
        {
            string vTextDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        vTextDateTimeShort = vDateTime.ToString("yyyy-MM-dd", null);
                    }
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vTextDateTimeShort;
        }

        #endregion;

        #region ----- Set OverTime ----- 

        private bool SET_OT_REQ_TIME(int pIDX_ROW, object pOT_HEADER_ID,  object pOT_FLAG, object pWORK_DATE, object pPERSON_ID)
        {
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDC_GET_OT_TIME.SetCommandParamValue("W_OT_HEADER_ID", pOT_HEADER_ID);
            IDC_GET_OT_TIME.SetCommandParamValue("W_OT_FLAG", pOT_FLAG);
            IDC_GET_OT_TIME.SetCommandParamValue("W_WORK_DATE", pWORK_DATE);
            IDC_GET_OT_TIME.SetCommandParamValue("W_PERSON_ID", pPERSON_ID);
            IDC_GET_OT_TIME.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_GET_OT_TIME.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_GET_OT_TIME.GetCommandParamValue("O_MESSAGE"));
            if (IDC_GET_OT_TIME.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }

            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("BEFORE_OT_START"), IDC_GET_OT_TIME.GetCommandParamValue("O_BF_START_TIME"));
            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("BEFORE_OT_END"), IDC_GET_OT_TIME.GetCommandParamValue("O_BF_END_TIME"));
            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_START"), IDC_GET_OT_TIME.GetCommandParamValue("O_START_DATE"));
            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_START"), IDC_GET_OT_TIME.GetCommandParamValue("O_START_TIME"));
            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_END"), IDC_GET_OT_TIME.GetCommandParamValue("O_END_DATE"));
            igrOT_LINE.SetCellValue(pIDX_ROW, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_END"), IDC_GET_OT_TIME.GetCommandParamValue("O_END_TIME"));

            return true;
        }

        #endregion

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();

                    misDelete = false;
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                    if (idaOT_LINE.IsFocused)
                    {
                        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10346"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        //object vObject = REQ_TYPE.EditValue; //N : 정상, A : 추가

                        System.Windows.Forms.SendKeys.Send("{TAB}");
                        if (isOT_Line_Check() != false)
                        {
                            idaOT_LINE.AddOver();
                            igrOT_LINE.SetCellValue("WORK_DATE", WORK_DATE.EditValue);
                        }
                    }
                    else
                    {
                        if (isOT_Header_Check() != false)
                        {
                            idaOT_HEADER.AddOver();
                            isOT_Header();
                        }
                    }
                    
                    misDelete = false;
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaOT_LINE.IsFocused)
                    {
                        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10346"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        //object vObject = REQ_TYPE.EditValue; //N : 정상, A : 추가

                        System.Windows.Forms.SendKeys.Send("{TAB}");
                        if (isOT_Line_Check() != false)
                        {
                            idaOT_LINE.AddUnder();
                            igrOT_LINE.SetCellValue("WORK_DATE", WORK_DATE.EditValue);
                        }
                    }
                    else
                    {                        
                        if (isOT_Header_Check() != false)
                        {
                            idaOT_HEADER.AddUnder();
                            isOT_Header();
                        }
                    }

                    misDelete = false;
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    System.Windows.Forms.SendKeys.Send("{TAB}");

                    try
                    {
                        if (misDelete == false)
                        {
                            bool vIsEqual = EqualWorkDate();
                            if (vIsEqual == false)
                            {
                                //[FCM_10393]신청하시는 연장근무, 모든 행 중, 근무일자가 동일 하지 않습니다. - 2011-10-17
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10393"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            if (isOT_Line_Check() == false)
                            {
                                return;
                            }

                            idaOT_HEADER.Update();
                        }
                        else
                        {
                            idaOT_HEADER.Update();

                            misDelete = false;
                        }
                    }
                    catch
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaOT_HEADER.IsFocused)
                    {
                        idaOT_LINE.Cancel();
                        idaOT_HEADER.Cancel();
                    }
                    else if (idaOT_LINE.IsFocused)
                    {
                        idaOT_LINE.Cancel();
                    }

                    misDelete = false;
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaOT_LINE.IsFocused)
                    {
                        idaOT_LINE.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print) //인쇄버튼
                {
                    XLPrinting1("PRINT", igrOT_LINE);
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export) //엑셀파일 버튼
                {
                    XLPrinting1("FILE", igrOT_LINE);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0380_Load(object sender, EventArgs e)
        {
            STD_DATE_0.EditValue = DateTime.Today;

            idaOT_HEADER.FillSchema();
            
            // Lookup SETTING
            ildREQ_TYPE.SetLookupParamValue("W_GROUP_CODE", "REQ_TYPE");
            ildREQ_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
            GetCapacity();
        }

        private void HRMF0380_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        #endregion;

        #region ----- Button Event -----

        #region ----- Event Remark -----

        //private void ibtSELECT_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        //{//대상산출
        //    int mRECORD_COUNT = 0;

        //    if (isOT_Line_Check() == false)
        //    {
        //        return;
        //    }

        //    idcOT_LINE_COUNT.ExecuteNonQuery();
        //    mRECORD_COUNT = Convert.ToInt32(idcOT_LINE_COUNT.GetCommandParamValue("O_RECORD_COUNT"));
        //    if (mRECORD_COUNT != Convert.ToInt32(0))
        //    {
        //        ////&&VALUE 는(은) 이미 존재합니다. &&TEXT 하시기 바랍니다.
        //        //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10044", "&&VALUE:=Request Number's Data(신청번호에 대한 라인자료)&&TEXT:=Search(조회)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        //return;

        //        //[2011-07-25]
        //        idaOT_HEADER.Cancel();
        //        //기준일자에 대한 연장근무 신청이 이미 존재 합니다. 신청No로 조회해 수정 하십시오!
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10301"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    WORK_DATE.EditValue = STD_DATE_0.EditValue;

        //    idaOT_LINE.Cancel();

        //    idaINSERT_PERSON.Fill();

        //    igrOT_LINE.BeginUpdate();
        //    for (int i = 0; i < idaINSERT_PERSON.OraDataSet().Rows.Count; i++)
        //    {
        //        idaOT_LINE.AddUnder();
        //        for (int j = 0; j < igrOT_LINE.GridAdvExColElement.Count - 1; j++)
        //        {
        //            igrOT_LINE.SetCellValue(i, j + 1, idaINSERT_PERSON.OraDataSet().Rows[i][j]);
        //        }
        //    }
        //    igrOT_LINE.EndUpdate();
        //    igrOT_LINE.CurrentCellMoveTo(0, 0);
        //    igrOT_LINE.CurrentCellActivate(0, 0);
        //    igrOT_LINE.Focus();
        //}

        #endregion;

        //[2011-10-27] - 대상산출
        private void ibtSELECT_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int mRECORD_COUNT = 0;

            if (isOT_Line_Check() == false)
            {
                return;
            }

            if (WORK_DATE.EditValue != null)
            {
                if (WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Saturday
                 || WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Sunday)
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                }
                else
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 1; //수정 가능, 조기출근 시작시간
                }
            }

            idcOT_LINE_COUNT.ExecuteNonQuery();
            mRECORD_COUNT = Convert.ToInt32(idcOT_LINE_COUNT.GetCommandParamValue("O_RECORD_COUNT"));
            if (mRECORD_COUNT != Convert.ToInt32(0))
            {
                //[2011-07-25]
                idaOT_HEADER.Cancel();
                //기준일자에 대한 연장근무 신청이 이미 존재 합니다. 신청No로 조회해 수정 하십시오!
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10301"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ProgressBar_ObjectWorkOut.Visible = true;

                WORK_DATE.EditValue = STD_DATE_0.EditValue;

                idaOT_LINE.Cancel();

                idaINSERT_PERSON.Fill();

                int vCountRow = idaINSERT_PERSON.OraSelectData.Rows.Count;
                int vCountColumn = idaINSERT_PERSON.OraSelectData.Columns.Count - 2;

                if (vCountRow > 0)
                {
                    igrOT_LINE.BeginUpdate();
                    for (int vROW = 0; vROW < vCountRow; vROW++)
                    {
                        idaOT_LINE.AddUnder();
                        for (int vCOL = 0; vCOL < vCountColumn; vCOL++)
                        {
                            igrOT_LINE.SetCellValue(vROW, vCOL, idaINSERT_PERSON.OraSelectData.Rows[vROW][vCOL]);
                        }

                        float vBarFill = ((float)vROW / (float)(vCountRow - 1)) * 100;
                        ProgressBar_ObjectWorkOut.BarFillPercent = vBarFill;
                    }
                    igrOT_LINE.EndUpdate();


                    object vObject_HOLY_TYPE_1 = null;
                    string vString_HOLY_TYPE_1 = string.Empty;

                    object vObject_HOLY_TYPE_2 = null;
                    string vString_HOLY_TYPE_2 = string.Empty;

                    int vIndexColumn_HOLY_TYPE_1 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_1");
                    int vIndexColumn_HOLY_TYPE_2 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_2");
                    int vIndexColumn_ALL_NIGHT_YN = igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN");

                    for (int vROW = 0; vROW < vCountRow; vROW++)
                    {
                        vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue(vROW, vIndexColumn_HOLY_TYPE_1);
                        vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                        vObject_HOLY_TYPE_2 = igrOT_LINE.GetCellValue(vROW, vIndexColumn_HOLY_TYPE_2);
                        vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

                        //0:무급유일[토], 1:휴일[일]
                        //주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화
                        if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                        {
                            //근무전연장 제어
                            igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                            igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
                            //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                            //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간

                            //근무후 일자
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 1; //수정 가능
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 1;
                            //근무후 시간
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 1; //수정 가능
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 1;

                            if (vString_HOLY_TYPE_2 == "3") //야간
                            {
                                igrOT_LINE.SetCellValue(vROW, vIndexColumn_ALL_NIGHT_YN, "Y");
                                SettingAllNight(igrOT_LINE, vROW);
                            }
                        }
                        else
                        {
                            //근무전연장 제어
                            igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 1; //수정 불가능, 조기출근 시작시간
                            igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 1; //수정 불가능, 조기출근 시작시간
                            //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Insertable = 1; //수정 불가능, 조기출근 시작시간
                            //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Updatable = 1; //수정 불가능, 조기출근 시작시간


                            //평일에는 근무후 시작일, 시작시를 수정할 일이 없어 수정 불가능하도록 설정
                            //근무후 일자
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0; //수정 불가능
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
                            //근무후 시간
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능
                            igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;
                        }
                    }
                }
                igrOT_LINE.CurrentCellMoveTo(0, 0);
                igrOT_LINE.CurrentCellActivate(0, 0);
                igrOT_LINE.Focus();

                ProgressBar_ObjectWorkOut.Visible = false;
            }
            catch (System.Exception ex)
            {
                ProgressBar_ObjectWorkOut.Visible = false;

                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void btnAPPR_REQUEST_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            bool vIsEqual = EqualWorkDate();
            if (vIsEqual == false)
            {
                //[FCM_10393]신청하시는 연장근무, 모든 행 중, 근무일자가 동일 하지 않습니다. - 2011-10-17
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10393"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (isOT_Line_Check() == false)
            {
                return;
            }

            //헤더 변경 데이터 존재 여부 체크 
            // 변경데이터 존재시 저장부터 하도록 유도 
            int vChg_Record_Count = 0;
            foreach (System.Data.DataRow vRow in idaOT_HEADER.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Unchanged)
                {
                    vChg_Record_Count++;
                }
            }
            if (vChg_Record_Count > 0)
            {
                //해당 자료는 이미 승인되었습니다. &&VALUE 할 수 없습니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //라인 변경 데이터 존재 여부 체크 
            // 변경데이터 존재시 저장부터 하도록 유도 
            vChg_Record_Count = 0;
            foreach (System.Data.DataRow vRow in idaOT_LINE.CurrentRows)
            {
                if (vRow.RowState != DataRowState.Unchanged)
                {
                    vChg_Record_Count++;
                }
            }
            if (vChg_Record_Count > 0)
            {
                //해당 자료는 이미 승인되었습니다. &&VALUE 할 수 없습니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            idaOT_HEADER.Update();

            if (Set_Request("OK") == false)
            {
                return;
            }           

            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "OT");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", REQ_DATE.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", REQ_DATE.EditValue);
            idcEMAIL_SEND.ExecuteNonQuery();

            // 다시 조회.
            //idaOT_HEADER.SetSelectParamValue("W_OT_HEADER_ID", OT_HEADER_ID_0.EditValue); 
            idaOT_HEADER.Fill();

            //승인요청을 하셨습니다.
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10277"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnAPPR_REQUEST_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Set_Request("CANCEL") == false)
            {
                return;
            }

            // 다시 조회.
            //idaOT_HEADER.SetSelectParamValue("W_OT_HEADER_ID", OT_HEADER_ID_0.EditValue); 
            idaOT_HEADER.Fill();
        }

        private void BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaOT_LINE.Cancel();

            igrOT_LINE.BeginUpdate();

            igrOT_LINE.CurrentCellMoveTo(0, 1);
            for (int r = 0; r < igrOT_LINE.RowCount; r++)
            {
                igrOT_LINE.CurrentCellMoveTo(r, 1);
                idaOT_LINE.Delete();
            }
            igrOT_LINE.EndUpdate();

            idaOT_HEADER.Delete();
            idaOT_HEADER.Update();
        }

        #endregion;

        #region ----- Grid Event -----

        private void igrOT_LINE_CellMoved(object pSender, ISGridAdvExCellClickEventArgs e)
        {
            int vIndexColumn_ALL_NIGHT_YN = igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN");

            if (e.ColIndex == vIndexColumn_ALL_NIGHT_YN)
            {
                System.Windows.Forms.SendKeys.Send("{TAB}");
            }
        }

        private void igrOT_LINE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == igrOT_LINE.GetColumnToIndex("WORK_DATE"))
            {
                if (e.NewValue != null)
                {
                    idcOT_STD_TIME_2.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                    idcOT_STD_TIME_2.SetCommandParamValue("W_WORK_DATE", e.NewValue);
                    idcOT_STD_TIME_2.SetCommandParamValue("W_DANGJIK_YN", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("DANGJIK_YN")));
                    idcOT_STD_TIME_2.SetCommandParamValue("W_ALL_NIGHT_YN", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN")));
                    idcOT_STD_TIME_2.ExecuteNonQuery();

                    igrOT_LINE.SetCellValue("BEFORE_OT_START", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_START"));
                    igrOT_LINE.SetCellValue("BEFORE_OT_END", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_END"));
                    igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_START"));
                    igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_START"));
                    igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                    igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_END"));
                }
            }
            else if (e.ColIndex == igrOT_LINE.GetColumnToIndex("DANGJIK_YN"))
            {
                idcOT_STD_TIME_2.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                idcOT_STD_TIME_2.SetCommandParamValue("W_WORK_DATE", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("WORK_DATE")));
                idcOT_STD_TIME_2.SetCommandParamValue("W_DANGJIK_YN", e.NewValue);
                idcOT_STD_TIME_2.SetCommandParamValue("W_ALL_NIGHT_YN", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN")));
                idcOT_STD_TIME_2.ExecuteNonQuery();

                igrOT_LINE.SetCellValue("BEFORE_OT_START", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_START"));
                igrOT_LINE.SetCellValue("BEFORE_OT_END", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_END"));
                igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_START"));
                igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_START"));
                igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_END"));
            }
            else if (e.ColIndex == igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN"))
            {
                object vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("HOLY_TYPE_1"));
                string vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                //HOLY_TYPE = 0 : 토
                //HOLY_TYPE = 1 : 일
                if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 1;  //수정 가능, 근무후 일자
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 1; //수정 가능, 근무후 시간

                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 1;
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 1;

                    idcOT_STD_TIME_2.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                    idcOT_STD_TIME_2.SetCommandParamValue("W_WORK_DATE", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("WORK_DATE")));
                    idcOT_STD_TIME_2.SetCommandParamValue("W_DANGJIK_YN", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("DANGJIK_YN")));
                    idcOT_STD_TIME_2.SetCommandParamValue("W_ALL_NIGHT_YN", e.NewValue);
                    idcOT_STD_TIME_2.ExecuteNonQuery();

                    igrOT_LINE.SetCellValue("BEFORE_OT_START", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_START"));
                    igrOT_LINE.SetCellValue("BEFORE_OT_END", idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_END"));
                    igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_START"));
                    igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_START"));
                    igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                    igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_END"));

                    object vObject_101 = igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("PERSON_ID"));
                    object vObject_102 = igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("WORK_DATE"));
                    object vObject_103 = igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("DANGJIK_YN"));
                    object vObject_104 = e.NewValue;

                    object vObject_105 = idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_START");
                    object vObject_106 = idcOT_STD_TIME_2.GetCommandParamValue("O_BEFORE_OT_END");
                    object vObject_107 = idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_START");
                    object vObject_108 = idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_START");
                    object vObject_109 = idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_END");
                    object vObject_110 = idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_END");
                }
                else
                {
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0;  //수정 불가능, 근무후 일자
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능, 근무후 시간

                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
                    igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;
                    
                    if (vString_HOLY_TYPE_1 == "2") //주간
                    {
                        idcOT_STD_TIME_2.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                        idcOT_STD_TIME_2.SetCommandParamValue("W_WORK_DATE", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("WORK_DATE")));
                        idcOT_STD_TIME_2.SetCommandParamValue("W_DANGJIK_YN", igrOT_LINE.GetCellValue(e.RowIndex, igrOT_LINE.GetColumnToIndex("DANGJIK_YN")));
                        idcOT_STD_TIME_2.SetCommandParamValue("W_ALL_NIGHT_YN", e.NewValue);
                        idcOT_STD_TIME_2.ExecuteNonQuery();

                        igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                        igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", idcOT_STD_TIME_2.GetCommandParamValue("O_AFTER_OT_TIME_END"));
                    }
                    else if (vString_HOLY_TYPE_1 == "3") //야간
                    {
                        igrOT_LINE.SetCellValue(e.RowIndex, e.ColIndex, "N");
                    }
                }
            }
        }
        
        //----------------------------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------------------------
        private void igrOT_LINE_CellDoubleClick(object pSender)
        {
            if (igrOT_LINE.RowIndex < 0 && igrOT_LINE.GetColumnToIndex("OT_FLAG") == igrOT_LINE.ColIndex)
            {
                object vOT_FLAG = "N";
                int vIDX_OT_FLAG = igrOT_LINE.GetColumnToIndex("OT_FLAG");
                int vIDX_OT_HEADER_ID = igrOT_LINE.GetColumnToIndex("OT_HEADER_ID");
                int vIDX_WORK_DATE = igrOT_LINE.GetColumnToIndex("WORK_DATE");
                int vIDX_PERSON_ID = igrOT_LINE.GetColumnToIndex("PERSON_ID");

                for (int r = 0; r < igrOT_LINE.RowCount; r++)
                {
                    if (iString.ISNull(igrOT_LINE.GetCellValue(r, vIDX_OT_FLAG), "N") == "Y".ToString())
                    {
                        vOT_FLAG = "N";
                    }
                    else
                    {
                        vOT_FLAG = "Y";                        
                    }
                    igrOT_LINE.SetCellValue(r, vIDX_OT_FLAG, vOT_FLAG);
                    bool vOT_REQ_STATUS = SET_OT_REQ_TIME(r, igrOT_LINE.GetCellValue(r, vIDX_OT_HEADER_ID), 
                                                        vOT_FLAG, igrOT_LINE.GetCellValue(r, vIDX_WORK_DATE), 
                                                        igrOT_LINE.GetCellValue(r, vIDX_PERSON_ID));
                    if (vOT_REQ_STATUS == false)
                    {
                        return;
                    }
                }
                
            }
            else if (igrOT_LINE.RowIndex < 0 && igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN"))
            {
                int vCountRow = igrOT_LINE.RowCount;

                if (vCountRow > 0)
                {
                    object vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue(0, igrOT_LINE.GetColumnToIndex("HOLY_TYPE_1"));
                    string vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                    //HOLY_TYPE = 0 : 토
                    //HOLY_TYPE = 1 : 일
                    if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                    {
                        for (int r = 0; r < igrOT_LINE.RowCount; r++)
                        {
                            if (iString.ISNull(igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN")), "N") == "Y".ToString())
                            {
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN"), "N");
                                idcOT_STD_TIME_1.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                                idcOT_STD_TIME_1.SetCommandParamValue("W_WORK_DATE", igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("WORK_DATE")));
                                idcOT_STD_TIME_1.SetCommandParamValue("W_DANGJIK_YN", "N");
                                idcOT_STD_TIME_1.SetCommandParamValue("W_ALL_NIGHT_YN", "N");
                                idcOT_STD_TIME_1.ExecuteNonQuery();

                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_END"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("BEFORE_OT_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("BEFORE_OT_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_END"));


                                object vObject_01 = igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("PERSON_ID"));
                                object vObject_02 = igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("WORK_DATE"));

                                object vObject_04 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_START");
                                object vObject_05 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_START");
                                object vObject_06 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_END");
                                object vObject_07 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_END");
                                object vObject_08 = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_START");
                                object vObject_09 = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_END");
                            }
                            else
                            {
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN"), "Y");
                                idcOT_STD_TIME_1.SetCommandParamValue("W_PERSON_ID", igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("PERSON_ID")));
                                idcOT_STD_TIME_1.SetCommandParamValue("W_WORK_DATE", igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("WORK_DATE")));
                                idcOT_STD_TIME_1.SetCommandParamValue("W_DANGJIK_YN", igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("DANGJIK_YN")));
                                idcOT_STD_TIME_1.SetCommandParamValue("W_ALL_NIGHT_YN", "Y");
                                idcOT_STD_TIME_1.ExecuteNonQuery();

                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_END"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_END"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("BEFORE_OT_START"), idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_START"));
                                igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("BEFORE_OT_END"), idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_END"));


                                object vObject_01 = igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("PERSON_ID"));
                                object vObject_02 = igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("WORK_DATE"));

                                object vObject_04 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_START");
                                object vObject_05 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_START");
                                object vObject_06 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_END");
                                object vObject_07 = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_END");
                                object vObject_08 = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_START");
                                object vObject_09 = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_END");
                            }
                        }
                    }
                }
            }
            else if (igrOT_LINE.RowIndex < 0 && igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("LUNCH_YN"))
            {
                for (int r = 0; r < igrOT_LINE.RowCount; r++)
                {
                    if (iString.ISNull(igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("LUNCH_YN")), "N") == "Y".ToString())
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("LUNCH_YN"), "N");                        
                    }
                    else
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("LUNCH_YN"), "Y");
                    }
                }       
            }
            else if (igrOT_LINE.RowIndex < 0 && igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("DINNER_YN"))
            {
                for (int r = 0; r < igrOT_LINE.RowCount; r++)
                {
                    if (iString.ISNull(igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("DINNER_YN")), "N") == "Y".ToString())
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("DINNER_YN"), "N");
                    }
                    else
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("DINNER_YN"), "Y");
                    }
                }
            }
            else if (igrOT_LINE.RowIndex < 0 && igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("MIDNIGHT_YN"))
            {
                for (int r = 0; r < igrOT_LINE.RowCount; r++)
                {
                    if (iString.ISNull(igrOT_LINE.GetCellValue(r, igrOT_LINE.GetColumnToIndex("MIDNIGHT_YN")), "N") == "Y".ToString())
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("MIDNIGHT_YN"), "N");
                    }
                    else
                    {
                        igrOT_LINE.SetCellValue(r, igrOT_LINE.GetColumnToIndex("MIDNIGHT_YN"), "Y");
                    }
                }
            }
            else if (igrOT_LINE.RowIndex > -1 && igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_END"))
            {
                //근무후 연장 종료시각
                object vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue("HOLY_TYPE_1");
                string vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                object vObject_HOLY_TYPE_2 = igrOT_LINE.GetCellValue("HOLY_TYPE_2");
                string vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

                object vObject_AFTER_OT_DATE_END = igrOT_LINE.GetCellValue("AFTER_OT_DATE_END");
                object vObject_AFTER_OT_TIME_END = igrOT_LINE.GetCellValue("AFTER_OT_TIME_END");

                string vMessage = string.Format("{0} | {1} | {2} | {3}", vString_HOLY_TYPE_1, vString_HOLY_TYPE_2, vObject_AFTER_OT_DATE_END, vObject_AFTER_OT_TIME_END);
                isAppInterfaceAdv1.OnAppMessage(vMessage);
                System.Windows.Forms.Application.DoEvents();

                //0:무급유일[토], 1:휴일[일]
                if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                {

                    if (vString_HOLY_TYPE_2 == "2") //주간
                    {
                        if (mIsSwitch == false)
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_AFTER_OT_DATE_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "17:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);

                            mIsSwitch = true;
                        }
                        else
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_AFTER_OT_DATE_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "21:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);

                            mIsSwitch = false;
                        }
                    }
                    else if (vString_HOLY_TYPE_2 == "3") //야간
                    {
                        if (mIsSwitch == false)
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_AFTER_OT_DATE_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "06:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);

                            mIsSwitch = true;
                        }
                        else
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_AFTER_OT_DATE_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "08:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);

                            mIsSwitch = false;
                        }
                    }
                }
            }
            else if (igrOT_LINE.RowIndex > -1 && (igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_START") || igrOT_LINE.ColIndex == igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_START")))
            {
                //근무후 연장 시작일시
                object vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue("HOLY_TYPE_1");
                string vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

                object vObject_HOLY_TYPE_2 = igrOT_LINE.GetCellValue("HOLY_TYPE_2");
                string vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

                object vObject_AFTER_OT_DATE_START = igrOT_LINE.GetCellValue("AFTER_OT_DATE_START");
                object vObject_AFTER_OT_TIME_START = igrOT_LINE.GetCellValue("AFTER_OT_TIME_START");
                object vObject_AFTER_OT_DATE_END = null;
                object vObject_AFTER_OT_TIME_END = null;

                object vObject_WORK_DATE = igrOT_LINE.GetCellValue("WORK_DATE");

                string vMessage = string.Format("{0} | {1} | {2} | {3}", vString_HOLY_TYPE_1, vString_HOLY_TYPE_2, vObject_AFTER_OT_DATE_START, vObject_AFTER_OT_TIME_START);
                isAppInterfaceAdv1.OnAppMessage(vMessage);
                System.Windows.Forms.Application.DoEvents();

                if (vObject_AFTER_OT_DATE_START == null || vObject_AFTER_OT_TIME_START == null)
                {
                    //0:무급유일[토], 1:휴일[일]
                    if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
                    {
                        if (vString_HOLY_TYPE_2 == "2") //주간
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_WORK_DATE);
                            //vObject_AFTER_OT_START = new System.DateTime(vDateTime.Year, vDateTime.Month, vDateTime.Day, 08, 30, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_START", vObject_AFTER_OT_START);
                            vObject_AFTER_OT_DATE_START = vDateTime;
                            vObject_AFTER_OT_TIME_START = "08:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", vObject_AFTER_OT_DATE_START);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", vObject_AFTER_OT_TIME_START);

                            //vObject_AFTER_OT_END = new System.DateTime(vDateTime.Year, vDateTime.Month, vDateTime.Day, 21, 00, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_END", vObject_AFTER_OT_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "21:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);
                        }
                        else if (vString_HOLY_TYPE_2 == "3") //야간
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_WORK_DATE);
                            //vObject_AFTER_OT_START = new System.DateTime(vDateTime.Year, vDateTime.Month, vDateTime.Day, 21, 00, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_START", vObject_AFTER_OT_START);
                            vObject_AFTER_OT_DATE_START = vDateTime;
                            vObject_AFTER_OT_TIME_START = "21:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", vObject_AFTER_OT_DATE_START);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", vObject_AFTER_OT_TIME_START);

                            //vObject_AFTER_OT_END = new System.DateTime(vDateTime.Year, vDateTime.Month, (vDateTime.Day + 1), 08, 30, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_END", vObject_AFTER_OT_END);
                            vObject_AFTER_OT_DATE_END = vDateTime.AddDays(1);
                            vObject_AFTER_OT_TIME_END = "08:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);
                        }
                    }
                    else if (vString_HOLY_TYPE_1 == "2" || vString_HOLY_TYPE_1 == "3") //2:주, 3:야
                    {
                        if (vString_HOLY_TYPE_2 == "2") //주간
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_WORK_DATE);
                            //vObject_AFTER_OT_START = new System.DateTime(vDateTime.Year, vDateTime.Month, vDateTime.Day, 18, 00, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_START", vObject_AFTER_OT_START);
                            vObject_AFTER_OT_DATE_START = vDateTime;
                            vObject_AFTER_OT_TIME_START = "18:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", vObject_AFTER_OT_DATE_START);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", vObject_AFTER_OT_TIME_START);

                            //vObject_AFTER_OT_END = new System.DateTime(vDateTime.Year, vDateTime.Month, vDateTime.Day, 21, 00, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_END", vObject_AFTER_OT_END);
                            vObject_AFTER_OT_DATE_END = vDateTime;
                            vObject_AFTER_OT_TIME_END = "21:00";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);
                        }
                        else if (vString_HOLY_TYPE_2 == "3") //야간
                        {
                            System.DateTime vDateTime = ConvertDateTime(vObject_WORK_DATE);
                            //vObject_AFTER_OT_START = new System.DateTime(vDateTime.Year, vDateTime.Month, (vDateTime.Day + 1), 06, 30, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_START", vObject_AFTER_OT_START);
                            vObject_AFTER_OT_DATE_START = vDateTime.AddDays(1);
                            vObject_AFTER_OT_TIME_START = "06:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_START", vObject_AFTER_OT_DATE_START);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_START", vObject_AFTER_OT_TIME_START);

                            //vObject_AFTER_OT_END = new System.DateTime(vDateTime.Year, vDateTime.Month, (vDateTime.Day + 1), 08, 30, 00);
                            //igrOT_LINE.SetCellValue("AFTER_OT_END", vObject_AFTER_OT_END);
                            vObject_AFTER_OT_DATE_END = vDateTime.AddDays(1);
                            vObject_AFTER_OT_TIME_END = "08:30";
                            igrOT_LINE.SetCellValue("AFTER_OT_DATE_END", vObject_AFTER_OT_DATE_END);
                            igrOT_LINE.SetCellValue("AFTER_OT_TIME_END", vObject_AFTER_OT_TIME_END);
                        }
                    }
                }
            }
        }

        private void igrOT_LINE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_OT_FLAG = igrOT_LINE.GetColumnToIndex("OT_FLAG");
            if (e.ColIndex == vIDX_OT_FLAG)
            {
                bool vOT_REQ_FLAG = SET_OT_REQ_TIME(igrOT_LINE.RowIndex, igrOT_LINE.GetCellValue("OT_HEADER_ID"),
                                                    e.NewValue, igrOT_LINE.GetCellValue("WORK_DATE"), igrOT_LINE.GetCellValue("PERSON_ID"));

            }
        }

        private void igrOT_LINE_CellKeyDown(object pSender, KeyEventArgs e)
        {
            //사용하지 않는 코드
            //int vIndexRowCurrent = igrOT_LINE.RowIndex;
            //int vIndexColumnCurrent = igrOT_LINE.ColIndex;

            //int vIndexColumn_BEFORE_OT_START = igrOT_LINE.GetColumnToIndex("BEFORE_OT_START");
            //int vIndexColumn_BEFORE_OT_END = igrOT_LINE.GetColumnToIndex("BEFORE_OT_END");

            //if (vIndexColumnCurrent == vIndexColumn_BEFORE_OT_START)
            //{
            //    if (e.KeyCode == System.Windows.Forms.Keys.Delete)
            //    {
            //        object vObject_DB_NULL = System.DBNull.Value;

            //        igrOT_LINE.SetCellValue(vIndexRowCurrent, vIndexColumn_BEFORE_OT_START, vObject_DB_NULL);
            //        igrOT_LINE.SetCellValue(vIndexRowCurrent, vIndexColumn_BEFORE_OT_END, vObject_DB_NULL);
            //    }
            //}
        }

        private void igrOT_LINE_CurrentCellAcceptedChanges(object pSender, ISGridAdvExChangedEventArgs e)
        {
            //사용자가 근무후 종료일자를 삭제 했다면, 근무후 시작일자도 지우도록 함.
            int vIndexColumn_AFTER_OT_DATE_START = igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_START");
            int vIndexColumn_AFTER_OT_TIME_START = igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_START");
            int vIndexColumn_AFTER_OT_DATE_END = igrOT_LINE.GetColumnToIndex("AFTER_OT_DATE_END");
            int vIndexColumn_AFTER_OT_TIME_END = igrOT_LINE.GetColumnToIndex("AFTER_OT_TIME_END");

            if (e.ColIndex == vIndexColumn_AFTER_OT_DATE_END || e.ColIndex == vIndexColumn_AFTER_OT_TIME_END)
            {
                object vObject = e.NewValue;
                if (vObject == null || iString.ISNull(vObject) == string.Empty)
                {
                    object vObject_DB_NULL = System.DBNull.Value;

                    igrOT_LINE.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_DATE_START, vObject_DB_NULL);
                    igrOT_LINE.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_TIME_START, vObject_DB_NULL);

                    igrOT_LINE.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_DATE_END, vObject_DB_NULL);
                    igrOT_LINE.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_TIME_END, vObject_DB_NULL);
                }
            }
        }

        //----------------------------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------------------------------------------

        #endregion

        #region ----- Adapter Event ------

        private void idaOT_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["REQ_TYPE"]) == string.Empty)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Type(신청구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["REQ_DATE"] == DBNull.Value)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Date(신청 일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DUTY_MANAGER_ID"] == DBNull.Value)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Control Level(근태관리 단위)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["REQ_PERSON_ID"] == DBNull.Value)
            {
                //&&VALUE는(은) 필수입니다. 확인하세요
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Person(신청자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaOT_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            //if (igrOT_LINE.RowCount != 0)
            //{// 라인 존재.
            //    //라인내역이 존재하므로 헤더를 삭제할 수 없습니다. 라인내역을 모두 삭제 후 헤더를 삭제해 주시기 바랍니다.
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}

            //if (e.Row["OT_HEADER_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Request Number(신청 번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        private void idaOT_LINE_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }

            string vHOLY_TYPE_1 = iString.ISNull(pBindingManager.DataRow["HOLY_TYPE_1"]);

            //0:무급유일[토], 1:휴일[일]
            //주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화
            if (vHOLY_TYPE_1 == "0" || vHOLY_TYPE_1 == "1")
            {
                //근무전연장 제어
                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].LookupAdapter = null;
                igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].LookupAdapter = null;

                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
                //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Insertable = 0; //수정 불가능, 조기출근 시작시간
                //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
            }
            else
            {
                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].LookupAdapter = ilaSTART_TIME_BEFORE;
                igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].LookupAdapter = ilaEND_TIME_BEFORE;

                //근무전연장 제어
                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Insertable = 1; //수정 불가능, 조기출근 시작시간
                igrOT_LINE.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 1; //수정 불가능, 조기출근 시작시간
                //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Insertable = 1; //수정 불가능, 조기출근 시작시간
                //igrOT_LINE.GridAdvExColElement[mIDX_BF_END_TIME].Updatable = 1; //수정 불가능, 조기출근 시작시간
            }
        }

        private void idaOT_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if ((iString.ISNull(e.Row["HOLY_TYPE_1"]) == "2" && iString.ISNull(e.Row["ALL_NIGHT_YN"]) != "Y") &&
                (iString.ISNull(e.Row["BREAKFAST_YN"]) == "Y" || iString.ISNull(e.Row["MIDNIGHT_YN"]) == "Y"))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10510"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (iString.ISNull(e.Row["HOLY_TYPE_1"]) == "3" &&
                (iString.ISNull(e.Row["LUNCH_YN"]) == "Y" || iString.ISNull(e.Row["DINNER_YN"]) == "Y"))
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10511"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void idaOT_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10047"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        #endregion

        #region ----- LookUP Event ----

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
        }
      
        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
        }

        private void ilaDUTY_MANAGER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDUTY_MANAGER.SetLookupParamValue("W_END_DATE", STD_DATE_0.EditValue);
        }

        private void ilaPERSON_SelectedRowData(object pSender)
        {
            System.Windows.Forms.SendKeys.Send("{TAB}");
        }

        #endregion

        #region ----- Edit Event ----

        private void STD_DATE_0_EditValueChanged(object pSender)
        {
            if (idaOT_HEADER.CurrentRow == null)
            {
                return;
            }
            else if (idaOT_HEADER.CurrentRow.RowState == DataRowState.Added)
            {
                WORK_DATE.EditValue = STD_DATE_0.EditValue;
            }
        }

        #endregion

        #region ----- WorkDate Equal Method -----

        private bool EqualWorkDate()
        {
            bool vIsEqual = true; //true이면 모든 행이 같은 근무일자이며, false 이면 모든 행중 하나라도 틀린 근무일자 존재.
            int vCountFalse = 0;

            int vCountRow = igrOT_LINE.RowCount;
            int vIndexColumn = igrOT_LINE.GetColumnToIndex("WORK_DATE");

            object vObject_Edit = WORK_DATE.EditValue;
            object vObject_Grid = null;

            string vStringDate_Edit = ConvertDate(vObject_Edit);
            string vStringDate_Grid = string.Empty;

            for (int vRow = 0; vRow < vCountRow; vRow++)
            {
                vObject_Grid = igrOT_LINE.GetCellValue(vRow, vIndexColumn);
                vStringDate_Grid = ConvertDate(vObject_Grid);

                if (vStringDate_Edit != vStringDate_Grid)
                {
                    vCountFalse++;
                }
            }

            if (vCountFalse > 0)
            {
                vIsEqual = false;
            }

            return vIsEqual;
        }

        #endregion;

        #region ----- Setting All Night Method -----

        private void SettingAllNight(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pROW)
        {
            int vIndexColumn_PERSON_ID = pGrid.GetColumnToIndex("PERSON_ID");
            object vObject_PERSON_ID = pGrid.GetCellValue(pROW, vIndexColumn_PERSON_ID);

            int vIndexColumn_WORK_DATE = pGrid.GetColumnToIndex("WORK_DATE");
            object vObject_WORK_DATE = pGrid.GetCellValue(pROW, vIndexColumn_WORK_DATE);

            int vIndexColumn_DANGJIK_YN = pGrid.GetColumnToIndex("DANGJIK_YN");
            object vObject_DANGJIK_YN = pGrid.GetCellValue(pROW, vIndexColumn_DANGJIK_YN);

            idcOT_STD_TIME_1.SetCommandParamValue("W_PERSON_ID", vObject_PERSON_ID);
            idcOT_STD_TIME_1.SetCommandParamValue("W_WORK_DATE", vObject_WORK_DATE);
            idcOT_STD_TIME_1.SetCommandParamValue("W_DANGJIK_YN", vObject_DANGJIK_YN);
            idcOT_STD_TIME_1.SetCommandParamValue("W_ALL_NIGHT_YN", "Y");
            idcOT_STD_TIME_1.ExecuteNonQuery();


            int vIndexColumn_AFTER_OT_DATE_START = pGrid.GetColumnToIndex("AFTER_OT_DATE_START");
            int vIndexColumn_AFTER_OT_TIME_START = pGrid.GetColumnToIndex("AFTER_OT_TIME_START");
            int vIndexColumn_AFTER_OT_DATE_END = pGrid.GetColumnToIndex("AFTER_OT_DATE_END");
            int vIndexColumn_AFTER_OT_TIME_END = pGrid.GetColumnToIndex("AFTER_OT_TIME_END");

            int vIndexColumn_BEFORE_OT_START = pGrid.GetColumnToIndex("BEFORE_OT_START");
            int vIndexColumn_BEFORE_OT_END = pGrid.GetColumnToIndex("BEFORE_OT_END");

            object vObject_AFTER_OT_DATE_START = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_START");
            object vObject_AFTER_OT_TIME_START = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_START");
            object vObject_AFTER_OT_DATE_END = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_DATE_END");
            object vObject_AFTER_OT_TIME_END = idcOT_STD_TIME_1.GetCommandParamValue("O_AFTER_OT_TIME_END");

            object vObject_BEFORE_OT_START = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_START");
            object vObject_BEFORE_OT_END = idcOT_STD_TIME_1.GetCommandParamValue("O_BEFORE_OT_END");

            pGrid.SetCellValue(pROW, vIndexColumn_AFTER_OT_DATE_START, vObject_AFTER_OT_DATE_START);
            pGrid.SetCellValue(pROW, vIndexColumn_AFTER_OT_TIME_START, vObject_AFTER_OT_TIME_START);
            pGrid.SetCellValue(pROW, vIndexColumn_AFTER_OT_DATE_END, vObject_AFTER_OT_DATE_END);
            pGrid.SetCellValue(pROW, vIndexColumn_AFTER_OT_TIME_END, vObject_AFTER_OT_TIME_END);

            pGrid.SetCellValue(pROW, vIndexColumn_BEFORE_OT_START, vObject_BEFORE_OT_START);
            pGrid.SetCellValue(pROW, vIndexColumn_BEFORE_OT_END, vObject_BEFORE_OT_END);
        }

        #endregion;

        #region ----- ilaPERSON_SelectedRowData Setting All Night Method -----
        //현재 안 쓰고 있음.
        private void ilaPERSON_SelectedRowData_Method()
        {
            object vObject_HOLY_TYPE_1 = null;
            string vString_HOLY_TYPE_1 = string.Empty;

            object vObject_HOLY_TYPE_2 = null;
            string vString_HOLY_TYPE_2 = string.Empty;

            int vIndexColumn_HOLY_TYPE_1 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_1");
            int vIndexColumn_HOLY_TYPE_2 = igrOT_LINE.GetColumnToIndex("HOLY_TYPE_2");
            int vIndexColumn_ALL_NIGHT_YN = igrOT_LINE.GetColumnToIndex("ALL_NIGHT_YN");

            int vIndexRow = igrOT_LINE.RowIndex;

            vObject_HOLY_TYPE_1 = igrOT_LINE.GetCellValue(vIndexRow, vIndexColumn_HOLY_TYPE_1);
            vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

            vObject_HOLY_TYPE_2 = igrOT_LINE.GetCellValue(vIndexRow, vIndexColumn_HOLY_TYPE_2);
            vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

            //0:무급유일[토], 1:휴일[일]
            if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
            {
                if (vString_HOLY_TYPE_2 == "3") //야간
                {
                    igrOT_LINE.SetCellValue(vIndexRow, vIndexColumn_ALL_NIGHT_YN, "Y");
                    SettingAllNight(igrOT_LINE, vIndexRow);
                }
            }
        }

        #endregion;


        #region ----- Request Limit Method -----

        private bool RequestLimit()
        {
            string vString_REQUEST_LIMIT_COUNT = string.Empty; ;
            int vREQUEST_LIMIT_COUNT = 0;
            bool vIsRequest = true;

            if (mCAPACITY == "C")
            {
                return vIsRequest;
            }

            try
            {
                idcGET_REQUEST_LIMIT.SetCommandParamValue("W_CODE", "OT_LIMIT");
                idcGET_REQUEST_LIMIT.ExecuteNonQuery();
                object o_REQUEST_LIMIT_COUNT = idcGET_REQUEST_LIMIT.GetCommandParamValue("O_REQUEST_LIMIT_COUNT");
                vString_REQUEST_LIMIT_COUNT = ConvertString(o_REQUEST_LIMIT_COUNT);
                vREQUEST_LIMIT_COUNT = int.Parse(vString_REQUEST_LIMIT_COUNT);

                System.DateTime vWorkDate = WORK_DATE.DateTimeValue;

                System.DateTime vCurrentDate = GetDate();
                System.TimeSpan vSubtractDate = vCurrentDate - vWorkDate;

                int vSubDay = vSubtractDate.Days;
                if (vSubDay > vREQUEST_LIMIT_COUNT) //3
                {
                    vIsRequest = false;

                    //일 전의 연장근무를 신청할 수 없습니다!
                    string vMessage = string.Format("{0}{1}", vREQUEST_LIMIT_COUNT, isMessageAdapter1.ReturnText("FCM_10421"));
                    MessageBoxAdv.Show(vMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsRequest;
        }

        #endregion;

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        #endregion;

        #region ----- XL Print 1 Method ----

        private void XLPrinting1(string pOutChoice, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            int vCountRow = pGrid.RowCount;

            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Format("XL Open...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //string vUserLogin = string.Format("{0}[{1}]", isAppInterfaceAdv1.AppInterface.LoginDescription, isAppInterfaceAdv1.DEPT_NAME);

                string vREQ_TYPE_NAME = string.Format("{0}", REQ_TYPE_NAME.EditValue); //신청구분
                string vREQ_NUM = string.Format("신청번호 : {0}", REQ_NUM.EditValue); //신청번호
                string vDUTY_MANAGER_NAME = string.Format("작업장 : {0}", DUTY_MANAGER_NAME.EditValue); //작업장
                string vREQ_PERSON_NAME = string.Format("신청자 : {0}", REQ_PERSON_NAME.EditValue); //신청자
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0380_001.xls";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                if (isOpen == true)
                {
                    vPageNumber = xlPrinting.LineWrite(pGrid, vREQ_TYPE_NAME, vREQ_NUM, vDUTY_MANAGER_NAME, vREQ_PERSON_NAME);

                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.SAVE("OT_");
                    }

                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------

                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

    }
}