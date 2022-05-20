using InfoSummit.Win.ControlAdv;
using ISCommonUtil;
using Syncfusion.Windows.Forms;
using Syncfusion.XlsIO;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;


namespace HRMF0364
{
    public partial class HRMF0364 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();
         
        private bool mSave_Flag = false; 
        private string mCAPACITY = string.Empty;


        //그리드 col 제어위해 그리드 col index 값 정의 
        private int mIDX_BF_START_TIME = 8;   //근무전 연장 시시간 
        private int mIDX_BF_END_TIME = 9;   //근무전 연장 시시간 
        private int mIDX_AF_START_DATE = 10;   //근무전 연장 시시간 
        private int mIDX_AF_START_TIME = 11;   //근무전 연장 시시간 

        //object mSOURCE_CATEGORY = null;
        
        private ISFileTransferAdv mFileTransfer;
        private isFTP_Info mFTP_Info;

        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 실행 디렉토리.        
        private string mDownload_Folder = string.Empty;             // Download Folder 
        private bool mFTP_Connect_Status = false;                   // FTP 정보 상태.

        #endregion;

        #region ----- Constructor -----

        public HRMF0364(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            
            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)   //파견직관리
            { 
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A; 
            }
        }

        #endregion;

        #region ----- Corp Type -----

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

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
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y"); 

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W1_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W1_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
            W1_CORP_NAME.BringToFront();

            G_CORP_GROUP.BringToFront();
            G1_CORP_TYPE.BringToFront();
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iConv.ISNull(G_CORP_TYPE.EditValue, "1") == "1")
            {
                G_CORP_GROUP.Visible = false; //.Show();
                V_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G_CORP_TYPE.EditValue = V_RB_OWNER.RadioCheckedString;

                G1_CORP_GROUP.Visible = false;
                V1_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G1_CORP_TYPE.EditValue = V1_RB_OWNER.RadioCheckedString;
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

                G1_CORP_GROUP.Visible = true;
                if (iConv.ISNull(G1_CORP_TYPE.EditValue) == "ALL")
                {
                    V1_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G1_CORP_TYPE.EditValue = V1_RB_ALL.RadioCheckedString;
                }
                else
                {
                    V1_RB_ETC.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G1_CORP_TYPE.EditValue = V1_RB_ETC.RadioCheckedString;
                }
            }
        }

        private void GetCapacity()
        {
            try
            {
                IDC_GET_CAPACITY.ExecuteNonQuery();
                object oCAPACITY = IDC_GET_CAPACITY.GetCommandParamValue("O_CAPACITY_C");

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
            if (TB_MAIN.SelectedTab.TabIndex == TP_OT_LIST.TabIndex)
            {
                if (iConv.ISNull(W1_CORP_ID.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W1_CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_CORP_NAME.Focus();
                    return;
                }
                if (iConv.ISNull(W1_WORK_DATE_FR.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W1_WORK_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W1_WORK_DATE_FR.Focus();
                    return;
                }
                if (iConv.ISNull(W1_WORK_DATE_TO.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W1_WORK_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W1_WORK_DATE_TO.Focus();
                    return;
                } 
                 
                IDA_OT_REQ_LIST.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                IDA_OT_REQ_LIST.Fill();
                IGR_OT_REQ_LIST.Focus();
            }
            else 
            {
                if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_CORP_NAME.Focus();
                    return;
                }
                if (iConv.ISNull(W_WORK_DATE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_WORK_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_WORK_DATE.Focus();
                    return;
                } 
                if (W_WORK_DATE.EditValue != null)
                {
                    if (W_WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Saturday
                     || W_WORK_DATE.DateTimeValue.DayOfWeek == System.DayOfWeek.Sunday)
                    {
                        IGR_OT_REQ.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 0; //수정 불가능, 조기출근 시작시간
                    }
                    else
                    {
                        IGR_OT_REQ.GridAdvExColElement[mIDX_BF_START_TIME].Updatable = 1; //수정 가능, 조기출근 시작시간
                    }
                }

                IGR_OT_REQ.LastConfirmChanges();
                IDA_OT_REQ.OraSelectData.AcceptChanges();
                IDA_OT_REQ.Refillable = true;

                IDA_OT_REQ.SetSelectParamValue("W_SOB_ID", -1);
                IDA_OT_REQ.Fill();

                V_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_DANGJIK_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_ALL_NIGHT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_OT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

                ////주말, 휴일에는 언제 출근할지 몰라서, 근무후 시작일, 시작시를 수정 가능하도록 활성화
                ////평일에는 근무후 시작일, 시작시를 수정할 일이 없어 수정 불가능하도록 설정
                //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Insertable = 0;  //수정 불가능, 근무후 일자
                //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Insertable = 0; //수정 불가능, 근무후 시간

                //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_DATE].Updatable = 0;
                //igrOT_LINE.GridAdvExColElement[mIDX_AF_START_TIME].Updatable = 0;

                IDA_OT_REQ.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.AppInterface.SOB_ID);
                IDA_OT_REQ.Fill();
                IGR_OT_REQ.Focus();
            }
        }

        private void SEARCH_DB_Calendar(object pPerson_ID, object pWork_Date)
        { 
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_FR", iDate.ISDate_Add(pWork_Date, -3));
            IDA_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_TO", pWork_Date);
            IDA_WORK_CALENDAR_S.Fill(); 
        }

        private void SEARCH_DB_ATTACHMENT(object pSOURCE_CATEGORY, object pSOURCE_ID)
        {
            //이미지 초기화;
            //ImageView(string.Empty);

            //첨부파일 리스트 조회 
            IDA_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_CATEGORY", pSOURCE_CATEGORY);
            IDA_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_ID", pSOURCE_ID);
            IDA_DOC_ATTACHMENT.Fill();
        }

        private void DELETE_DOC_ATTACHMENT()
        {
            object vDOC_ATTACHMENT_ID = IGR_DOC_ATTACHMENT.GetCellValue("DOC_ATTACHMENT_ID");
            if (iConv.ISNull(vDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DeleteFile(vDOC_ATTACHMENT_ID) == false)
            {
                return;
            }
            
            SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_OT_REQ.GetCellValue("OT_ID"));
        }


        private void SEARCH_DB_Calendar1(object pPerson_ID, object pWork_Date)
        {
            IDA1_WORK_CALENDAR_S.SetSelectParamValue("W_PERSON_ID", pPerson_ID);
            IDA1_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_FR", iDate.ISDate_Add(pWork_Date, -3));
            IDA1_WORK_CALENDAR_S.SetSelectParamValue("W_WORK_DATE_TO", pWork_Date);
            IDA1_WORK_CALENDAR_S.Fill();
        }

        private void SEARCH_DB_ATTACHMENT1(object pSOURCE_CATEGORY, object pSOURCE_ID)
        {
            //이미지 초기화;
            //ImageView(string.Empty);

            //첨부파일 리스트 조회 
            IDA1_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_CATEGORY", pSOURCE_CATEGORY);
            IDA1_DOC_ATTACHMENT.SetSelectParamValue("P_SOURCE_ID", pSOURCE_ID);
            IDA1_DOC_ATTACHMENT.Fill();
        }

        private void Set_OT_STD_Time(int pRow_Index)
        {
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_PERSON_ID", IGR_OT_REQ.GetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("PERSON_ID")));
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_WORK_DATE", IGR_OT_REQ.GetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("WORK_DATE")));
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_HOLY_TYPE", IGR_OT_REQ.GetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("HOLY_TYPE")));
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_DANGJIK_YN", IGR_OT_REQ.GetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("DANGJIK_YN")));
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_ALL_NIGHT_YN", IGR_OT_REQ.GetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("ALL_NIGHT_YN")));
            IDC_GET_OT_STD_TIME.ExecuteNonQuery();
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_START"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_TIME_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_END"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_TIME_END"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_START"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_DATE_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_END"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_DATE_END"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_END"));

            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_START_M"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_TIME_START_M"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_END_M"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_TIME_END_M"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START_M"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_START_M"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END_M"), IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_END_M"));
        }

        private void Get_Req_OT_Time(int pRow_Index, object pWork_date, object pPerson_ID, object pHoly_Type, object pOT_Time_Type)
        {
            IDC_GET_REQ_OT_TIME_P.SetCommandParamValue("W_WORK_DATE", pWork_date);
            IDC_GET_REQ_OT_TIME_P.SetCommandParamValue("W_PERSON_ID", pPerson_ID);
            IDC_GET_REQ_OT_TIME_P.SetCommandParamValue("W_HOLY_TYPE", pHoly_Type);
            IDC_GET_REQ_OT_TIME_P.SetCommandParamValue("W_OT_TIME_TYPE", pOT_Time_Type);
            IDC_GET_REQ_OT_TIME_P.ExecuteNonQuery();  
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_START"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_BEFORE_OT_TIME_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_END"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_BEFORE_OT_TIME_END"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_START_M"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_BEFORE_OT_TIME_START_M"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_END_M"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_BEFORE_OT_TIME_END_M"));
            
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_START"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_DATE_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_TIME_START"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START_M"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_TIME_START_M"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_END"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_DATE_END"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_TIME_END"));
            IGR_OT_REQ.SetCellValue(pRow_Index, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END_M"), IDC_GET_REQ_OT_TIME_P.GetCommandParamValue("O_AFTER_OT_TIME_END_M"));
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

        private bool SAVE_CHECK_P(object pOT_ID, object pWORK_DATE, object pPERSON_ID, object pDANGJIK_YN, object pALL_NIGHT_YN
                                , object pOT_FLAG
                                , object pBEFORE_TIME_START, object pBEFORE_TIME_START_M
                                , object pBEFORE_TIME_END, object pBEFORE_TIME_END_M
                                , object pAFTER_DATE_START, object pAFTER_TIME_START, object pAFTER_TIME_START_M
                                , object pAFTER_DATE_END, object pAFTER_TIME_END, object pAFTER_TIME_END_M)
        {
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_OT_ID", pOT_ID);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_WORK_DATE", pWORK_DATE);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_PERSON_ID", pPERSON_ID);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_DANGJIK_YN", pDANGJIK_YN);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_ALL_NIGHT_YN", pALL_NIGHT_YN);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_OT_FLAG", pOT_FLAG);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_BEFORE_TIME_START", pBEFORE_TIME_START);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_BEFORE_TIME_START_M", pBEFORE_TIME_START_M);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_BEFORE_TIME_END", pBEFORE_TIME_END);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_BEFORE_TIME_END_M", pBEFORE_TIME_END_M);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_DATE_START", pAFTER_DATE_START);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_TIME_START", pAFTER_TIME_START);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_TIME_START_M", pAFTER_TIME_START_M);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_DATE_END", pAFTER_DATE_END);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_TIME_END", pAFTER_TIME_END);
            IDC_SAVE_CHECK_P.SetCommandParamValue("W_AFTER_TIME_END_M", pAFTER_TIME_END_M); 

            IDC_SAVE_CHECK_P.ExecuteNonQuery();
            string vSTATUS = iConv.ISNull(IDC_SAVE_CHECK_P.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iConv.ISNull(IDC_SAVE_CHECK_P.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                if(vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return false;
            }

            return true;
        }

        #endregion;

        #region ----- FTP Infomation ----- 
        //ftp 접속정보 및 환경 정보 설정 
        private void Set_FTP_Info()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            mFTP_Connect_Status = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "HR_OT");
                IDC_FTP_INFO.ExecuteNonQuery();
                if (IDC_FTP_INFO.ExcuteError)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                    return;
                }

                mFTP_Info = new isFTP_Info();

                mFTP_Info.Host = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mFTP_Info.Port = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mFTP_Info.UserID = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mFTP_Info.Password = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mFTP_Info.Passive_Flag = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));
                mFTP_Info.FTP_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mFTP_Info.Client_Folder = iConv.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            if (mFTP_Info.Host == string.Empty)
            {
                //ftp접속정보 오류          
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            try
            {
                //FileTransfer Initialze
                mFileTransfer = new ISFileTransferAdv();
                mFileTransfer.Host = mFTP_Info.Host;
                mFileTransfer.Port = mFTP_Info.Port;
                mFileTransfer.UserId = mFTP_Info.UserID;
                mFileTransfer.Password = mFTP_Info.Password;
                if (mFTP_Info.Passive_Flag == "Y")
                {
                    mFileTransfer.UsePassive = true;
                }
                else
                {
                    mFileTransfer.UsePassive = false;
                }
                mDownload_Folder = string.Format("{0}\\{1}", mClient_Base_Path, mFTP_Info.Client_Folder);
            }
            catch (System.Exception Ex)
            {
                //ftp접속정보 오류 
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }

            //Client Download Folder 없으면 생성 
            System.IO.DirectoryInfo vDownload_Folder = new System.IO.DirectoryInfo(mDownload_Folder);
            if (vDownload_Folder.Exists == false) //있으면 True, 없으면 False
            {
                vDownload_Folder.Create();
            }

            mFTP_Connect_Status = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion

        #region ----- File Upload Methods -----
        //ftp에 file upload 처리 
        private bool UpLoadFile(object pDOC_REV_ID, object pDOCUMENT_REV_NUM)
        {
            bool isUpload = false;

            if (mFTP_Connect_Status == false)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iConv.ISNull(pDOCUMENT_REV_NUM) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "All File(*.*)|*.*|Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|PowerPoint File(*.ppt;*.pptx)|*.ppt;*.pptx|jpg file(*.jpg)|*.jpg|Pdf File(*.pdf)|*.pdf";
                openFileDialog1.DefaultExt = "*.*";
                openFileDialog1.FileName = "";
                openFileDialog1.Multiselect = false;

                //openFileDialog1.Title = "Select Open File";
                //openFileDialog1.Filter = "jpg file(*.jpg)|*.jpg|bmp file(*.bmp)|*.bmp";
                //openFileDialog1.DefaultExt = "jpg";
                //openFileDialog1.FileName = "";
                //openFileDialog1.Multiselect = false;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    //1. 사용자 선택 파일 
                    string vSelectFullPath = openFileDialog1.FileName;
                    string vSelectDirectoryPath = Path.GetDirectoryName(openFileDialog1.FileName);

                    string vFileName = Path.GetFileName(openFileDialog1.FileName);
                    string vFileExtension = Path.GetExtension(openFileDialog1.FileName).ToUpper();

                    //transaction 이용하기 위해 설정
                    IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = isDataTransaction1;
                    IDC_INSERT_DOC_ATTACHMENT.DataTransaction = isDataTransaction1;

                    //2. 첨부파일 DB 저장
                    isDataTransaction1.BeginTran();

                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_CATEGORY", V_DOC_CATEGORY.EditValue); //구분 
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_SOURCE_ID", pDOC_REV_ID);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_FTP_FILE_NAME", pDOCUMENT_REV_NUM);
                    IDC_INSERT_DOC_ATTACHMENT.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                    IDC_INSERT_DOC_ATTACHMENT.ExecuteNonQuery();

                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));
                    object vDOC_ATTACHMENT_ID = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_DOC_ATTACHMENT_ID");
                    object vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACHMENT.GetCommandParamValue("O_FTP_FILE_NAME");

                    //O_DOC_ATTACHMENT_ID.EditValue = vDOC_ATTACHMENT_ID;
                    //O_FTP_FILE_NAME.EditValue = vFTP_FILE_NAME;

                    if (IDC_INSERT_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        isDataTransaction1.RollBack();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        //Transaction 해제.
                        IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                        IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                        return isUpload;
                    }

                    //3. 첨부파일 로그 저장 
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
                    IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                    IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        isDataTransaction1.RollBack();
                        if (vMESSAGE != string.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        //Transaction 해제.
                        IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                        IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                        return isUpload;
                    }

                    //4. 파일 업로드
                    try
                    {
                        int vArryCount = openFileDialog1.FileNames.Length;
                        for (int r = 0; r < vArryCount; r++)
                        {
                            mFileTransfer.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransfer.SourceDirectory = vSelectDirectoryPath;
                            mFileTransfer.SourceFileName = vFileName;
                            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
                            mFileTransfer.TargetFileName = iConv.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransfer.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false;
                                isDataTransaction1.RollBack();
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                //Transaction 해제.
                                IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                                return isUpload;
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        isDataTransaction1.RollBack();
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                        return isUpload;
                    }

                    //5. 적용
                    isDataTransaction1.Commit();
                    //Transaction 해제.
                    IDC_INSERT_DOC_ATTACHMENT.DataTransaction = null;
                    IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                }
            }
            return isUpload;
        }

        #endregion;

        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(string pFileName)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();            
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", O_DOC_ATTACHMENT_ID.EditValue);
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            //IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            //vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            //if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            //{
            //    Application.UseWaitCursor = false;
            //    this.Cursor = Cursors.Default;
            //    Application.DoEvents();

            //    isDataTransaction1.RollBack();
            //    if (vMESSAGE != string.Empty)
            //    {
            //        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //    return IsDownload;
            //}            

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFileName);
            string vClientFileName = string.Format("{0}", pFileName);

            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.SourceFileName = pFileName;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시
                    string vClientFullPath = string.Format("{0}\\{1}", mDownload_Folder, vClientFileName);  //원본

                    System.IO.File.Delete(vClientFullPath);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, vClientFullPath);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        //ftp file download 처리 
        private bool DownLoadFile(string pSAVE_FileName, string pFTP_FILE_NAME)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_FILE_ENTRY_ID", pFILE_ENTRY_ID);
            //IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            //IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            //vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            //vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            //if (vSTATUS == "F")
            //{
            //    isDataTransaction1.RollBack();
            //    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFTP_FILE_NAME);
            string vClientFileName = string.Format("{0}", pSAVE_FileName);

            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.SourceFileName = pFTP_FILE_NAME;
            mFileTransfer.TargetDirectory = mDownload_Folder;
            mFileTransfer.TargetFileName = vTempFileName;

            IsDownload = mFileTransfer.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mDownload_Folder, vTempFileName);      //임시
                    string vClientFullPath = string.Format("{0}", vClientFileName);  //원본

                    System.IO.File.Delete(vClientFullPath);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, vClientFullPath);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }

            //isDataTransaction1.Commit();
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;

        #region ----- is View file Method -----

        private string isDownload(object pFileName)
        {
            string vFileName = iConv.ISNull(pFileName);

            if (vFileName != string.Empty)
            {
                if (DownLoadFile(vFileName) == true)
                {
                    return string.Format("{0}\\{1}", mDownload_Folder, vFileName);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        private string isDownload(string pSAVE_FileName, string pFTP_FILE_NAME)
        {
            if (pSAVE_FileName != string.Empty && pFTP_FILE_NAME != string.Empty)
            {
                if (DownLoadFile(pSAVE_FileName, pFTP_FILE_NAME) == true)
                {
                    return string.Format("{0}", pSAVE_FileName);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

        #endregion;


        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(object pDOC_ATTACHMENT_ID)
        {
            bool IsDelete = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            object vDOC_ATTACHMENT_ID = pDOC_ATTACHMENT_ID;
            string vFTP_FileName = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            if (iConv.ISNull(vDOC_ATTACHMENT_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }
            if (iConv.ISNull(vFTP_FileName) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            //transaction 이용하기 위해 설정
            IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = isDataTransaction1;
            IDC_DELETE_DOC_ATTACHMENT.DataTransaction = isDataTransaction1;

            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            isDataTransaction1.BeginTran();
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
            IDC_INSERT_DOC_ATTACHMENT_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "DELETE");
            IDC_INSERT_DOC_ATTACHMENT_LOG.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_INSERT_DOC_ATTACHMENT_LOG.GetCommandParamValue("O_MESSAGE"));
            if (IDC_INSERT_DOC_ATTACHMENT_LOG.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                isDataTransaction1.RollBack();
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }

            //2. 파일 삭제 
            IDC_DELETE_DOC_ATTACHMENT.SetCommandParamValue("W_DOC_ATTACHMENT_ID", vDOC_ATTACHMENT_ID);
            IDC_DELETE_DOC_ATTACHMENT.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_DELETE_DOC_ATTACHMENT.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACHMENT.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false;
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }

            //3. 실제 삭제  
            mFileTransfer.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransfer.SourceDirectory = mFTP_Info.FTP_Folder;  //삭제는 소스에 설정해야 삭제됨.
            mFileTransfer.SourceFileName = vFTP_FileName;
            mFileTransfer.TargetDirectory = mFTP_Info.FTP_Folder;
            mFileTransfer.TargetFileName = vFTP_FileName;

            IsDelete = mFileTransfer.Delete();
            if (IsDelete == false)
            {
                isDataTransaction1.RollBack();
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                //Transaction 해제.
                IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
                IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
                return IsDelete;
            }
            isDataTransaction1.Commit();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            //Transaction 해제.
            IDC_DELETE_DOC_ATTACHMENT.DataTransaction = null;
            IDC_INSERT_DOC_ATTACHMENT_LOG.DataTransaction = null;
            return IsDelete;
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

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Get Date Method -----

        private bool Set_Request(string pStatus)
        {
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;


            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("SELECT_FLAG");
            int vIDX_OT_ID = IGR_OT_REQ.GetColumnToIndex("OT_ID");
            int vIDX_PERSON_ID = IGR_OT_REQ.GetColumnToIndex("PERSON_ID");
            int vIDX_WORK_DATE = IGR_OT_REQ.GetColumnToIndex("WORK_DATE");

            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_OT_REQ.GetCellValue(r, vIDX_SELECT_FLAG)))
                {
                    IDC_SET_UPDATE_REQUEST.SetCommandParamValue("P_OT_ID", IGR_OT_REQ.GetCellValue(r, vIDX_OT_ID));
                    IDC_SET_UPDATE_REQUEST.SetCommandParamValue("P_WORK_DATE", IGR_OT_REQ.GetCellValue(r, vIDX_WORK_DATE));
                    IDC_SET_UPDATE_REQUEST.SetCommandParamValue("P_PERSON_ID", IGR_OT_REQ.GetCellValue(r, vIDX_PERSON_ID));
                    IDC_SET_UPDATE_REQUEST.SetCommandParamValue("P_REQUEST_STATUS", pStatus);
                    IDC_SET_UPDATE_REQUEST.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_UPDATE_REQUEST.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_REQUEST.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_UPDATE_REQUEST.ExcuteError || vSTATUS == "F")
                    {
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

        #endregion;

        #region ----- Get Date Method -----

        private DateTime GetDate()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                IDC_GET_DATE.ExecuteNonQuery();
                object vObject = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

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

        #region ----- Excel Export II -----

        private void ExcelExport(ISDataAdapter pAdapter, ISGridAdvEx pGrid)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.RestoreDirectory = true;

            //기본 저장 경로 지정.            
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            vSaveFileName = "Person List";     //기본 파일명. 수정필요.

            saveFileDialog1.Title = "Excel Save";
            saveFileDialog1.FileName = vSaveFileName;
            saveFileDialog1.Filter = "CSV File(*.csv)|*.csv|Excel file(*.xlsx)|*.xlsx|Excel file(*.xls)|*.xls";
            saveFileDialog1.DefaultExt = "xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                vSaveFileName = saveFileDialog1.FileName;
                System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                try
                {
                    if (vFileName.Exists)
                    {
                        vFileName.Delete();
                    }
                }
                catch (Exception EX)
                {
                    MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            vMessageText = string.Format(" Writing Starting...");

            System.Windows.Forms.Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //DATA 조회   
            int vCountRow = pAdapter.CurrentRows.Count;

            if (vCountRow < 1)
            {
                vMessageText = isMessageAdapter1.ReturnText("EAPP_10106");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            try
            {
                //Step 1 : Instantiate the spreadsheet creation engine.
                ExcelEngine ExcelEngine = new ExcelEngine();

                //Step 2 : Instantiate the excel application object.
                IApplication Exc_App = ExcelEngine.Excel;

                //set 2.1 : file Extension check =>xlsx, xls 
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel97to2003;
                }
                else
                {
                    ExcelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;
                }

                //A new workbook is created.[Equivalent to creating a new workbook in MS Excel]
                //The new workbook will have 3 worksheets
                IWorkbook Exc_WorkBook = Exc_App.Workbooks.Create(1);
                if (Path.GetExtension(vSaveFileName).ToUpper() == ".XLS")
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel97to2003;
                }
                else
                {
                    Exc_WorkBook.Version = ExcelVersion.Excel2007;
                }

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet sheet = Exc_WorkBook.Worksheets[0];

                //Export DataTable.
                sheet.ImportDataTable(pAdapter.OraDataTable(), false, 1, 1, pAdapter.CurrentRows.Count, pAdapter.OraSelectData.Columns.Count, true);

                //1.title insert
                int vHeaderCount = pGrid.GridAdvExColElement[0].HeaderElement.Count;
                for (int h = 1; h <= vHeaderCount; h++)
                {
                    sheet.InsertRow(h);
                    object vTitle = string.Empty;
                    for (int c = 0; c < pGrid.ColCount; c++)
                    {
                        if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL1_KR)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL1_KR;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL2_CN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL2_CN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL3_VN)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL3_VN;
                        }
                        else if (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage == ISUtil.Enum.TerritoryLanguage.TL4_JP)
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].TL4_JP;
                        }
                        else
                        {
                            vTitle = pGrid.GridAdvExColElement[c].HeaderElement[vHeaderCount - h].Default;
                        }

                        sheet.Range[h, c + 1].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                        sheet.Range[h, c + 1].Value = iConv.ISNull(vTitle);
                        sheet.AutofitColumn(c + 1);
                        if (iConv.ISNull(pGrid.GridAdvExColElement[c].Visible) == "0")
                        {
                            sheet.HideColumn(c + 1);
                        }
                    }
                }

                ////2.prompt insert
                //sheet.InsertRow(2);
                //sheet.ImportDataTable(IDA_REJECT_DETAIL_TITLE.OraDataTable(), false, 2, 1); 
                //Exc_WorkBook.ActiveSheet.AutofitColumn(1);

                //Saving the workbook to disk.
                Exc_WorkBook.SaveAs(vSaveFileName);

                //Close the workbook.
                Exc_WorkBook.Close();

                //No exception will be thrown if there are unsaved workbooks.
                ExcelEngine.ThrowNotSavedOnDestroy = false;
                ExcelEngine.Dispose();

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(vSaveFileName);
                }

            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                System.Windows.Forms.Application.DoEvents();
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;


        #region ----- MDi ToolBar Button Event -----

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
                    if (IDA_OT_REQ.IsFocused)
                    {
                        IDA_OT_REQ.AddOver();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_OT_REQ.IsFocused)
                    {
                        IDA_OT_REQ.AddUnder();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    System.Windows.Forms.SendKeys.Send("{TAB}");
                    try
                    {
                        bool vIsEqual = EqualWorkDate();
                        if (vIsEqual == false)
                        {
                            //[FCM_10393]신청하시는 연장근무, 모든 행 중, 근무일자가 동일 하지 않습니다. - 2011-10-17
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10393"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        //if (isOT_Line_Check() == false)
                        //{
                        //    return;
                        //}
                        IDA_OT_REQ.Update(); 
                      
                    }
                    catch(Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_OT_REQ.IsFocused)
                    {
                        IDA_OT_REQ.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_OT_REQ.IsFocused)
                    {
                        IDA_OT_REQ.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print) //인쇄버튼
                {
                    XLPrinting1("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export) //엑셀파일 버튼
                {
                    ExcelExport(IDA_OT_REQ_LIST, IGR_OT_REQ_LIST);
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0364_Load(object sender, EventArgs e)
        {
            W1_WORK_DATE_FR.EditValue = iDate.ISDate_Add(DateTime.Today, -3);
            W1_WORK_DATE_TO.EditValue = DateTime.Today;

            W_WORK_DATE.EditValue = DateTime.Today;
            IDA_OT_REQ.FillSchema();

            // LOOKUP DEFAULT VALUE SETTING - DUTY_APPROVE_STATUS
            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            W_APPROVAL_STATUS_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME");
            W_APPROVAL_STATUS.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE");

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
            GetCapacity();
            Set_FTP_Info();
             
            IDC_GET_PERSON_NAME_P.SetCommandParamValue("P_PERSON_ID", isAppInterfaceAdv1.AppInterface.PersonId);
            IDC_GET_PERSON_NAME_P.ExecuteNonQuery();
            V_REQ_NAME.EditValue = IDC_GET_PERSON_NAME_P.GetCommandParamValue("O_PERSON_NAME");
            V_REQ_PERSON_NUM.EditValue = IDC_GET_PERSON_NAME_P.GetCommandParamValue("O_PERSON_NUM"); 
        }

        private void V_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("SELECT_FLAG"); 
            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                IGR_OT_REQ.SetCellValue(r, vIDX_SELECT_FLAG, V_SELECT_YN.CheckBoxString);
            } 
        }

        private void CB_DANGJIK_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("DANGJIK_YN");
            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                IGR_OT_REQ.SetCellValue(r, vIDX_SELECT_FLAG, CB_DANGJIK_YN.CheckBoxString);
                Set_OT_STD_Time(r);
            }
        }

        private void CB_ALL_NIGHT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("ALL_NIGHT_YN");
            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                IGR_OT_REQ.SetCellValue(r, vIDX_SELECT_FLAG, CB_ALL_NIGHT_YN.CheckBoxString);
                Set_OT_STD_Time(r);
            }
        }

        private void CB_OT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("OT_FLAG");
            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                IGR_OT_REQ.SetCellValue(r, vIDX_SELECT_FLAG, CB_OT_YN.CheckBoxString);
            }
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
        private void BTN_GET_PERSON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int mRECORD_COUNT = 0;

            if (W_WORK_DATE.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_WORK_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_OT_COUNT_DATA.ExecuteNonQuery();
            mRECORD_COUNT = Convert.ToInt32(IDC_OT_COUNT_DATA.GetCommandParamValue("O_RECORD_COUNT"));
            //if (mRECORD_COUNT != Convert.ToInt32(0))
            //{
            //    //[2011-07-25]
            //    idaOT_HEADER.Cancel();
            //    //기준일자에 대한 연장근무 신청이 이미 존재 합니다. 신청No로 조회해 수정 하십시오!
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10301"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            try
            {
                PB_GET_PERSON.Visible = true; 
                IDA_OT_REQ.Cancel();

                CB_DANGJIK_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_ALL_NIGHT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
                CB_OT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

                IDA_INSERT_DATA.SetSelectParamValue("W_LOOKUP_YN", "N");
                IDA_INSERT_DATA.Fill();

                int vCountRow = IDA_INSERT_DATA.OraSelectData.Rows.Count;
                int vCountColumn = IDA_INSERT_DATA.OraSelectData.Columns.Count - 2;

                IDA_OT_REQ.MoveLast(IGR_OT_REQ.Name);
                int vIDX_CURR = IGR_OT_REQ.RowIndex;
                if(vIDX_CURR == -1)
                {
                    IDA_OT_REQ.Cancel(); 
                }
                vIDX_CURR = IGR_OT_REQ.RowIndex;

                if (vCountRow > 0)
                {
                    IGR_OT_REQ.BeginUpdate();
                    for (int vROW = 0; vROW < vCountRow; vROW++)
                    {
                        IDA_OT_REQ.AddUnder();
                        for (int vCOL = 0; vCOL < vCountColumn; vCOL++)
                        {
                            IGR_OT_REQ.SetCellValue(vROW + (vIDX_CURR + 1), vCOL, IDA_INSERT_DATA.OraSelectData.Rows[vROW][vCOL]);
                        }

                        float vBarFill = ((float)vROW / (float)(vCountRow - 1)) * 100;
                        PB_GET_PERSON.BarFillPercent = vBarFill;
                    }
                    IGR_OT_REQ.EndUpdate(); 
                }
                IGR_OT_REQ.CurrentCellMoveTo(0, 0);
                IGR_OT_REQ.CurrentCellActivate(0, 0);
                IGR_OT_REQ.Focus();

                PB_GET_PERSON.Visible = false;
            }
            catch (System.Exception ex)
            {
                PB_GET_PERSON.Visible = false;

                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void BTN_SELECT_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vIDX_SELECT_FLAG = IGR_OT_REQ.GetColumnToIndex("SELECT_FLAG");
            for (int r = 0; r < IGR_OT_REQ.RowCount; r++)
            {
                if ("Y" == iConv.ISNull(IGR_OT_REQ.GetCellValue(r, vIDX_SELECT_FLAG)))
                {
                    IGR_OT_REQ.CurrentCellActivate(r, vIDX_SELECT_FLAG);
                    IGR_OT_REQ.CurrentCellMoveTo(r, vIDX_SELECT_FLAG);

                    IDA_OT_REQ.Delete();
                }
            }
            IGR_OT_REQ.LastConfirmChanges();
        }

        private void BTN_APPR_REQ_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            bool vIsEqual = EqualWorkDate();
            if (vIsEqual == false)
            {
                //[FCM_10393]신청하시는 연장근무, 모든 행 중, 근무일자가 동일 하지 않습니다. - 2011-10-17
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10393"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //변경된 자료 존재 여부 확인// 
            mSave_Flag = true;
            IDA_OT_REQ.Update();
            if (mSave_Flag == false)
            {
                return;
            }

            decimal vCnt = 0;
            foreach(DataRow vRow in IDA_OT_REQ.CurrentRows)
            {
                if(vRow.RowState != DataRowState.Unchanged)
                { 
                    vCnt++;
                } 
            }
            if (vCnt > 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            //승인요청.
            if (Set_Request("OK") == false)
            {
                return;
            }

            // EMAIL 발송.
            try
            {
                IDC_EMAIL_SEND.SetCommandParamValue("P_GUBUN", "A");
                IDC_EMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "OT");
                IDC_EMAIL_SEND.SetCommandParamValue("P_CORP_ID", W_CORP_ID.EditValue);
                IDC_EMAIL_SEND.SetCommandParamValue("P_WORK_DATE", W_WORK_DATE.EditValue);
                IDC_EMAIL_SEND.SetCommandParamValue("P_REQ_DATE", W_WORK_DATE.EditValue);
                IDC_EMAIL_SEND.ExecuteNonQuery();
            }
            catch
            {
                //
            }
             
            // 다시 조회.
            SEARCH_DB(); 
        }

        private void BTN_APPR_REQ_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Set_Request("CANCEL") == false)
            {
                return;
            }

            // 다시 조회. 
            SEARCH_DB(); 
        }

        private void BTN_SELECT_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            try
            {
                IDA_OT_REQ.Update();
            }
            catch (Exception Ex)
            {
                isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                return;
            }

            object vOT_ID = IGR_OT_REQ.GetCellValue("OT_ID");

            //Document Revision Update.
            if (iConv.ISNull(vOT_ID) == string.Empty)
            {
                return;
            }

            IDC_GET_DOC_LAST_FLAG.SetCommandParamValue("W_OT_ID", IGR_OT_REQ.GetCellValue("OT_ID"));
            IDC_GET_DOC_LAST_FLAG.ExecuteNonQuery();
            string vLAST_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_FLAG.GetCommandParamValue("O_LAST_FLAG"));
            if (vLAST_FLAG == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            object vDOCUMENT_REV_NUM = string.Format("{0}_{1:yyyyMMdd}", IGR_OT_REQ.GetCellValue("PERSON_NUM"), IGR_OT_REQ.GetCellValue("WORK_DATE"));

            if (UpLoadFile(vOT_ID, vDOCUMENT_REV_NUM) == true)
            {
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, vOT_ID);
            }
           // SEARCH_DB();
        }

        private void BTN_FILE_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            ////업로드 가능자 아니면 최종버전만 다운로드 가능하도록 제어//
            //if (mMANAGER_FLAG != "Y")
            //{
            //    IDC_GET_DOC_LAST_REV_FLAG.ExecuteNonQuery();
            //    string vLAST_REV_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_REV_FLAG.GetCommandParamValue("O_LAST_REV_FLAG"));
            //    if (vLAST_REV_FLAG != "Y")
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10174"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}

            if (mFTP_Connect_Status == false)
            {
                MessageBoxAdv.Show("FTP IP is not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 저장될 Dialog 열기
            saveFileDialog1.Title = "Select Save Folder";
            saveFileDialog1.FileName = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = "C:\\";
            saveFileDialog1.Filter = "All file(*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string vSAVE_FILE_NAME = saveFileDialog1.FileName;
                string vFTP_FILE_NAME = iConv.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
                try
                {
                    isDownload(vSAVE_FILE_NAME, vFTP_FILE_NAME);
                }
                catch
                {
                    MessageBox.Show("Error : Could not read file from disk.");

                }

                System.Diagnostics.Process.Start(vSAVE_FILE_NAME);
            }
        }

        private void BTN_FILE_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            object vOT_ID = IGR_OT_REQ.GetCellValue("OT_ID");
            if (iConv.ISNull(vOT_ID) == string.Empty)
            {
                return;
            }

            IDC_GET_DOC_LAST_FLAG.SetCommandParamValue("W_OT_ID", IGR_OT_REQ.GetCellValue("OT_ID"));
            IDC_GET_DOC_LAST_FLAG.ExecuteNonQuery();
            string vLAST_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_FLAG.GetCommandParamValue("O_LAST_FLAG"));
            if (vLAST_FLAG == "N")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10262"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10168"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            DELETE_DOC_ATTACHMENT();
        }

        private void BTN1_FILE_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR1_DOC_ATTACHMENT.RowIndex < 0)
            {
                return;
            }

            ////업로드 가능자 아니면 최종버전만 다운로드 가능하도록 제어//
            //if (mMANAGER_FLAG != "Y")
            //{
            //    IDC_GET_DOC_LAST_REV_FLAG.ExecuteNonQuery();
            //    string vLAST_REV_FLAG = iConv.ISNull(IDC_GET_DOC_LAST_REV_FLAG.GetCommandParamValue("O_LAST_REV_FLAG"));
            //    if (vLAST_REV_FLAG != "Y")
            //    {
            //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10174"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}

            if (mFTP_Connect_Status == false)
            {
                MessageBoxAdv.Show("FTP IP is not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 저장될 Dialog 열기
            saveFileDialog1.Title = "Select Save Folder";
            saveFileDialog1.FileName = iConv.ISNull(IGR1_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = "C:\\";
            saveFileDialog1.Filter = "All file(*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string vSAVE_FILE_NAME = saveFileDialog1.FileName;
                string vFTP_FILE_NAME = iConv.ISNull(IGR1_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
                try
                {
                    isDownload(vSAVE_FILE_NAME, vFTP_FILE_NAME);
                }
                catch
                {
                    MessageBox.Show("Error : Could not read file from disk.");

                }

                System.Diagnostics.Process.Start(vSAVE_FILE_NAME);
            }
        }

        #endregion;

        #region ----- Grid Event -----

        private void igrOT_LINE_CellMoved(object pSender, ISGridAdvExCellClickEventArgs e)
        {
            int vIndexColumn_ALL_NIGHT_YN = IGR_OT_REQ.GetColumnToIndex("ALL_NIGHT_YN");

            if (e.ColIndex == vIndexColumn_ALL_NIGHT_YN)
            {
                System.Windows.Forms.SendKeys.Send("{TAB}");
            }
        }
        
        private void IGR_OT_REQ_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_DANGJIK_YN = IGR_OT_REQ.GetColumnToIndex("DANGJIK_YN");
            int vIDX_ALL_NIGHT_YN = IGR_OT_REQ.GetColumnToIndex("ALL_NIGHT_YN");
            int vIDX_OT_FLAG = IGR_OT_REQ.GetColumnToIndex("OT_FLAG");

            if (e.ColIndex == vIDX_DANGJIK_YN || e.ColIndex == vIDX_ALL_NIGHT_YN || e.ColIndex == vIDX_OT_FLAG)
            {
                Set_OT_STD_Time(e.RowIndex); 
            }
        }
         
        private void igrOT_LINE_CurrentCellAcceptedChanges(object pSender, ISGridAdvExChangedEventArgs e)
        {
            //사용자가 근무후 종료일자를 삭제 했다면, 근무후 시작일자도 지우도록 함.
            int vIndexColumn_AFTER_OT_DATE_START = IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_START");
            int vIndexColumn_AFTER_OT_TIME_START = IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START");
            int vIndexColumn_AFTER_OT_DATE_END = IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_END");
            int vIndexColumn_AFTER_OT_TIME_END = IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END");

            if (e.ColIndex == vIndexColumn_AFTER_OT_DATE_END || e.ColIndex == vIndexColumn_AFTER_OT_TIME_END)
            {
                object vObject = e.NewValue;
                if (vObject == null || iConv.ISNull(vObject) == string.Empty)
                {
                    object vObject_DB_NULL = System.DBNull.Value;

                    IGR_OT_REQ.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_DATE_START, vObject_DB_NULL);
                    IGR_OT_REQ.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_TIME_START, vObject_DB_NULL);

                    IGR_OT_REQ.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_DATE_END, vObject_DB_NULL);
                    IGR_OT_REQ.SetCellValue(e.RowIndex, vIndexColumn_AFTER_OT_TIME_END, vObject_DB_NULL);
                }
            }
        }

        #endregion

        #region ----- Adapter Event ------

        private void IDA_OT_REQ_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                SEARCH_DB_Calendar(0, iDate.ISGetDate("1900-01-01"));
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, -1);
            }
            else
            {
                SEARCH_DB_Calendar(pBindingManager.DataRow["PERSON_ID"], pBindingManager.DataRow["WORK_DATE"]);
                SEARCH_DB_ATTACHMENT(V_DOC_CATEGORY.EditValue, IGR_OT_REQ.GetCellValue("OT_ID"));
            }
        }

        private void IDA_OT_REQ_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {            
            if (SAVE_CHECK_P(e.Row["OT_ID"], e.Row["WORK_DATE"], e.Row["PERSON_ID"], e.Row["DANGJIK_YN"], e.Row["ALL_NIGHT_YN"]
                                , e.Row["OT_FLAG"]
                                , e.Row["BEFORE_TIME_START"], e.Row["BEFORE_TIME_START_M"]
                                , e.Row["BEFORE_TIME_END"], e.Row["BEFORE_TIME_END_M"]
                                , e.Row["AFTER_OT_DATE_START"], e.Row["AFTER_OT_TIME_START"], e.Row["AFTER_OT_TIME_START_M"]
                                , e.Row["AFTER_OT_DATE_END"], e.Row["AFTER_OT_TIME_END"], e.Row["AFTER_OT_TIME_END_M"]) == false)
            {
                e.Cancel = true;
                mSave_Flag = false;
                return;
            }

            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("NAME")); 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["WORK_DATE"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("WORK_DATE"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["HOLY_TYPE"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("HOLY_NAME"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["BEFORE_TIME_START"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_START"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["BEFORE_TIME_END"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("BEFORE_TIME_END"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["AFTER_OT_DATE_START"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_START"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["AFTER_OT_TIME_START"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_START"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["AFTER_OT_DATE_END"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_DATE_END"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["AFTER_OT_TIME_END"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("AFTER_OT_TIME_END"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(e.Row["SELECT_FLAG"]) == "Y" && iConv.ISNull(e.Row["DESCRIPTION"]) == string.Empty)
            {
                mSave_Flag = false;
                e.Cancel = true;
                object vPrompt = Get_Grid_Prompt(IGR_OT_REQ, IGR_OT_REQ.GetColumnToIndex("DESCRIPTION"));
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", vPrompt)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }             
        }

        private void IDA_OT_REQ_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        private void IDA_OT_REQ_LIST_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                SEARCH_DB_Calendar1(0, iDate.ISGetDate("1900-01-01"));
                SEARCH_DB_ATTACHMENT1(V_DOC_CATEGORY.EditValue, -1);
            }
            else
            {
                SEARCH_DB_Calendar1(pBindingManager.DataRow["PERSON_ID"], pBindingManager.DataRow["WORK_DATE"]);
                SEARCH_DB_ATTACHMENT1(V_DOC_CATEGORY.EditValue, IGR_OT_REQ.GetCellValue("OT_ID"));
            }
        }

        #endregion

        #region ----- LookUP Event ----

        private void ILA_WORK_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_HOLY_TYPE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_APPROVAL_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_PERSON_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_W.SetLookupParamValue("W_START_DATE", W_WORK_DATE.EditValue);
            ILD_PERSON_W.SetLookupParamValue("W_END_DATE", W_WORK_DATE.EditValue);
        }
        
        private void ILA_DUTY_MANAGER_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_MANAGER.SetLookupParamValue("W_END_DATE", W_WORK_DATE.EditValue);
            ILD_DUTY_MANAGER.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
            ILD_DUTY_MANAGER.SetLookupParamValue("W_CAP_CHECK_YN", "Y"); 
        }

        private void ilaPERSON_SelectedRowData(object pSender)
        {
            System.Windows.Forms.SendKeys.Send("{TAB}");
        }

        private void ILA_PERSON_S_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERSON_S.SetLookupParamValue("W_LOOKUP_YN", "Y");
        }

        private void ILA_HOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_HOLY_TYPE_SelectedRowData(object pSender)
        {
            Set_OT_STD_Time(IGR_OT_REQ.RowIndex);
        }

        private void ILA_HOLY_TYPE_S_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "HOLY_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_OT_TIME_TYPE_SelectedRowData(object pSender)
        {
            int vRow_Index = IGR_OT_REQ.RowIndex;
            object vWork_date = IGR_OT_REQ.GetCellValue("WORK_DATE");
            object vPerson_ID = IGR_OT_REQ.GetCellValue("PERSON_ID");
            object vHoly_Type = IGR_OT_REQ.GetCellValue("HOLY_TYPE");
            object vOT_Time_Type = IGR_OT_REQ.GetCellValue("OT_TIME_TYPE");

            Get_Req_OT_Time(vRow_Index, vWork_date, vPerson_ID, vHoly_Type, vOT_Time_Type);
        }

        private void ILA_START_TIME_BEFORE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD_TIME.SetLookupParamValue("W_BASE_TIME", IGR_OT_REQ.GetCellValue("BEFORE_TIME_START"));
            ILD_PERIOD_TIME.SetLookupParamValue("W_BEFORE_YN", "Y"); 
            ILD_PERIOD_TIME.SetLookupParamValue("W_STD_TIME", DBNull.Value); 
        }

        private void ILA_START_TIME_BEFORE_SelectedRowData(object pSender)
        {
            //IGR_OT_REQ.SetCellValue("BEFORE_TIME_END", IGR_OT_REQ.GetCellValue("BEFORE_TIME_START"));
        }

        private void ILA_END_TIME_BEFORE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD_TIME.SetLookupParamValue("W_BASE_TIME", IGR_OT_REQ.GetCellValue("BEFORE_TIME_END"));
            ILD_PERIOD_TIME.SetLookupParamValue("W_BEFORE_YN", "Y"); 
            ILD_PERIOD_TIME.SetLookupParamValue("W_STD_TIME", IGR_OT_REQ.GetCellValue("BEFORE_TIME_START"));
        }

        private void ILA_START_TIME_AFTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD_TIME.SetLookupParamValue("W_BASE_TIME", IGR_OT_REQ.GetCellValue("AFTER_OT_TIME_START"));
            ILD_PERIOD_TIME.SetLookupParamValue("W_BEFORE_YN", "N"); 
            ILD_PERIOD_TIME.SetLookupParamValue("W_STD_TIME", DBNull.Value);
        }

        private void ILA_END_TIME_AFTER_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_PERIOD_TIME.SetLookupParamValue("W_BASE_TIME", IGR_OT_REQ.GetCellValue("AFTER_OT_TIME_END"));
            ILD_PERIOD_TIME.SetLookupParamValue("W_BEFORE_YN", "N"); 
            ILD_PERIOD_TIME.SetLookupParamValue("W_STD_TIME", IGR_OT_REQ.GetCellValue("AFTER_OT_TIME_START"));
        }

        private void ILA_START_TIME_AFTER_SelectedRowData(object pSender)
        {
            //IGR_OT_REQ.SetCellValue("AFTER_OT_TIME_END", IGR_OT_REQ.GetCellValue("AFTER_OT_TIME_START"));
        }

        private void ILA_END_TIME_AFTER_SelectedRowData(object pSender)
        {
            DateTime vAFTER_DATE_START = iDate.ISGetDate(IGR_OT_REQ.GetCellValue("AFTER_OT_DATE_START"));
            int vADD_DAY = iConv.ISNumtoZero(IGR_OT_REQ.GetCellValue("ADD_DAY_DATE_END"));

            IGR_OT_REQ.SetCellValue("AFTER_OT_DATE_END", iDate.ISDate_Add(vAFTER_DATE_START, vADD_DAY));
        }

        private void ILA_BREAKFAST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N"); 
        }

        private void ILA_LUNCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_DINNER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_MIDNIGHT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "Y");
        }

        private void ILA_BREAKFAST_S_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_LUNCH_S_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_DINNER_S_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "N");
        }

        private void ILA_MIDNIGHT_S_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OT_FOOD.SetLookupParamValue("W_ENABLED_FLAG", "Y");
            ILD_OT_FOOD.SetLookupParamValue("W_BREAKFAST_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_LUNCH_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_DINNER_FLAG", "N");
            ILD_OT_FOOD.SetLookupParamValue("W_MIDNIGHT_FLAG", "Y");
        }

        private void ILA_WORK_TYPE_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_APPROVAL_STATUS_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_APPROVE_STATUS");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_PERSON_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_W.SetLookupParamValue("W_START_DATE", W1_WORK_DATE_FR.EditValue);
            ILD_PERSON_W.SetLookupParamValue("W_END_DATE", W1_WORK_DATE_TO.EditValue);
        }

        private void ILA_DUTY_MANAGER_W1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY_MANAGER.SetLookupParamValue("W_END_DATE", W1_WORK_DATE_TO.EditValue);
            ILD_DUTY_MANAGER.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
            ILD_DUTY_MANAGER.SetLookupParamValue("W_CAP_CHECK_YN", "Y");
        }

        #endregion

        #region ----- Edit Event ----

        private void STD_DATE_0_EditValueChanged(object pSender)
        {
            //if (idaOT_HEADER.CurrentRow == null)
            //{
            //    return;
            //}
            //else if (idaOT_HEADER.CurrentRow.RowState == DataRowState.Added)
            //{
            //    WORK_DATE.EditValue = W_WORK_DATE.EditValue;
            //}
        }

        #endregion

        #region ----- WorkDate Equal Method -----

        private bool EqualWorkDate()
        {
            bool vIsEqual = true; //true이면 모든 행이 같은 근무일자이며, false 이면 모든 행중 하나라도 틀린 근무일자 존재.
            //int vCountFalse = 0;

            //int vCountRow = IGR_OT_REQ.RowCount;
            //int vIndexColumn = IGR_OT_REQ.GetColumnToIndex("WORK_DATE");

            //object vObject_Edit = WORK_DATE.EditValue;
            //object vObject_Grid = null;

            //string vStringDate_Edit = ConvertDate(vObject_Edit);
            //string vStringDate_Grid = string.Empty;

            //for (int vRow = 0; vRow < vCountRow; vRow++)
            //{
            //    vObject_Grid = IGR_OT_REQ.GetCellValue(vRow, vIndexColumn);
            //    vStringDate_Grid = ConvertDate(vObject_Grid);

            //    if (vStringDate_Edit != vStringDate_Grid)
            //    {
            //        vCountFalse++;
            //    }
            //}

            //if (vCountFalse > 0)
            //{
            //    vIsEqual = false;
            //}

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

            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_PERSON_ID", vObject_PERSON_ID);
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_WORK_DATE", vObject_WORK_DATE);
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_DANGJIK_YN", vObject_DANGJIK_YN);
            IDC_GET_OT_STD_TIME.SetCommandParamValue("W_ALL_NIGHT_YN", "Y");
            IDC_GET_OT_STD_TIME.ExecuteNonQuery();


            int vIndexColumn_AFTER_OT_DATE_START = pGrid.GetColumnToIndex("AFTER_OT_DATE_START");
            int vIndexColumn_AFTER_OT_TIME_START = pGrid.GetColumnToIndex("AFTER_OT_TIME_START");
            int vIndexColumn_AFTER_OT_DATE_END = pGrid.GetColumnToIndex("AFTER_OT_DATE_END");
            int vIndexColumn_AFTER_OT_TIME_END = pGrid.GetColumnToIndex("AFTER_OT_TIME_END");

            int vIndexColumn_BEFORE_OT_START = pGrid.GetColumnToIndex("BEFORE_OT_START");
            int vIndexColumn_BEFORE_OT_END = pGrid.GetColumnToIndex("BEFORE_OT_END");

            object vObject_AFTER_OT_DATE_START = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_DATE_START");
            object vObject_AFTER_OT_TIME_START = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_START");
            object vObject_AFTER_OT_DATE_END = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_DATE_END");
            object vObject_AFTER_OT_TIME_END = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_AFTER_OT_TIME_END");

            object vObject_BEFORE_OT_START = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_OT_START");
            object vObject_BEFORE_OT_END = IDC_GET_OT_STD_TIME.GetCommandParamValue("O_BEFORE_OT_END");

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

            int vIndexColumn_HOLY_TYPE_1 = IGR_OT_REQ.GetColumnToIndex("HOLY_TYPE_1");
            int vIndexColumn_HOLY_TYPE_2 = IGR_OT_REQ.GetColumnToIndex("HOLY_TYPE_2");
            int vIndexColumn_ALL_NIGHT_YN = IGR_OT_REQ.GetColumnToIndex("ALL_NIGHT_YN");

            int vIndexRow = IGR_OT_REQ.RowIndex;

            vObject_HOLY_TYPE_1 = IGR_OT_REQ.GetCellValue(vIndexRow, vIndexColumn_HOLY_TYPE_1);
            vString_HOLY_TYPE_1 = ConvertString(vObject_HOLY_TYPE_1);

            vObject_HOLY_TYPE_2 = IGR_OT_REQ.GetCellValue(vIndexRow, vIndexColumn_HOLY_TYPE_2);
            vString_HOLY_TYPE_2 = ConvertString(vObject_HOLY_TYPE_2);

            //0:무급유일[토], 1:휴일[일]
            if (vString_HOLY_TYPE_1 == "0" || vString_HOLY_TYPE_1 == "1")
            {
                if (vString_HOLY_TYPE_2 == "3") //야간
                {
                    IGR_OT_REQ.SetCellValue(vIndexRow, vIndexColumn_ALL_NIGHT_YN, "Y");
                    SettingAllNight(IGR_OT_REQ, vIndexRow);
                }
            }
        }

        #endregion;

         
        #region ----- XL Print 1 Method ----

        private void XLPrinting1(string pOutChoice)
        {
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;

            //프린트 데이터 조회//
            IDA_PRINT_OT_REQ.Fill();
            int vCountRow = IDA_PRINT_OT_REQ.CurrentRows.Count;
             
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data...");
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

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1, isMessageAdapter1);

            try
            {
                vMessageText = string.Format("Printing File Open...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                string vREQ_PERSON_NAME = string.Format("신청자 : {0}", V_REQ_NAME.EditValue); //신청자
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0364_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                if (isOpen == true)
                {
                    //인쇄일자 
                    IDC_GET_DATE.ExecuteNonQuery();
                    object vLOCAL_DATE = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                    vPageNumber = xlPrinting.XLWirteMain(IDA_PRINT_OT_REQ, vLOCAL_DATE, vREQ_PERSON_NAME);

                    if (pOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (pOutChoice == "FILE")
                    {
                        xlPrinting.Save("OT_");
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


    #region ----- FTP 정보 위한 사용자 Class -----

    public class isFTP_Info
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = "N";
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public isFTP_Info()
        {

        }

        public isFTP_Info(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag, string pFTP_Folder, string pClient_Folder)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mPassive_Flag = pPassive_Flag;
            mFTP_Folder = pFTP_Folder;
            mClient_Folder = pClient_Folder;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string Passive_Flag
        {
            get
            {
                return mPassive_Flag;
            }
            set
            {
                mPassive_Flag = value;
            }
        }

        public string FTP_Folder
        {
            get
            {
                return mFTP_Folder;
            }
            set
            {
                mFTP_Folder = value;
            }
        }

        public string Client_Folder
        {
            get
            {
                return mClient_Folder;
            }
            set
            {
                mClient_Folder = value;
            }
        }

        #endregion;
    }

    #endregion

}