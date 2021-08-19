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
using System.IO;

using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0303
{
    public partial class HRMF0303_CREATE : Office2007Form
    {
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0303_CREATE(Form pMainForm, ISAppInterface pAppInterface, object pCorp_ID, object pWork_YYYYMM, object pCORP_TYPE)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }

            CORP_ID.EditValue = pCorp_ID;
            WORK_YYYYMM.EditValue = pWork_YYYYMM;
        }

        #region ----- Private / Method ----- 

        private void Search_DB(object pCreated_Method)
        {
            IDA_CALENDAR_SET.Cancel();
            if (iConv.ISNull(WORK_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(WORK_TYPE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Type(교대 유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_TYPE.Focus();
                return;
            }

            //생성대상에 근무유형 설정 
            V_WORK_TYPE.EditValue = WORK_TYPE.EditValue;

            //2015.03.30. J.LAKE 변경 -> 근무계획 생성방법 변경 --             
            //if (iString.ISNull(pCreated_Method) == string.Empty)
            //{
            //    pCreated_Method = "A";
            //}
            //2015.03.30. J.LAKE 변경 -> 근무계획 생성방법 변경 -- 
            pCreated_Method = "A";

            // 기적용일수 조회.
            IDC_PRE_WORK_DAY_P.SetCommandParamValue("W_CREATED_METHOD", pCreated_Method);
            IDC_PRE_WORK_DAY_P.ExecuteNonQuery();
            PRE_WORK_DAY.EditValue = IDC_PRE_WORK_DAY_P.GetCommandParamValue("O_DAY_COUNT");

            IDA_CALENDAR_SET.SetSelectParamValue("W_CREATED_METHOD", pCreated_Method);
            IDA_CALENDAR_SET.Fill();

            int mRecordCount = IDA_CALENDAR_SET.SelectRows.Count;
            if (mRecordCount == 0)
            {
                Init_Work_Plan_STD();
            }

            SEARCH_DB_PRE();
            SEARCH_DB_DETAIL();
        }

        private void SEARCH_DB_DETAIL()
        {
            IDA_CALENDAR_DETAIL_1.Cancel();
            IDA_CALENDAR_DETAIL_2.Cancel();

            IDA_CALENDAR_DETAIL_1.Fill();
            IDA_CALENDAR_DETAIL_2.Fill();

        }
        private void SEARCH_DB_PRE()
        {
            string vPRE_MONTH = iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(String.Format("{0}-01", WORK_YYYYMM.EditValue)), -1));
            IDA_PRE_CAL_DETAIL_1.SetSelectParamValue("W_WORK_PERIOD", vPRE_MONTH);
            IDA_PRE_CAL_DETAIL_2.SetSelectParamValue("W_WORK_PERIOD", vPRE_MONTH);
            IDA_PRE_CAL_DETAIL_1.Fill();
            IDA_PRE_CAL_DETAIL_2.Fill();
        }

        private void Show_Import(bool pView_Flag)
        {
            if(pView_Flag == true)
            {
                UPLOAD_FILE_PATH.EditValue = String.Empty;
                V_START_ROW.EditValue = 2;
                V_MESSAGE.PromptText = "";
                V_PB_INTERFACE.BarFillPercent = 0;
                Application.DoEvents();

                GB_IMPORT.BringToFront();
                GB_IMPORT.Visible = true;
            }
            else

            {                
                GB_IMPORT.Visible = false;
            } 
        }

        private void isSetCommonLookUpParameter(string P_GROUP_CODE, string P_ENABLED_FLAG)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_ENABLED_FLAG);
        }

        private void isSetCommonLookUpParameter(string P_GROUP_CODE, object P_WHERE, string P_ENABLED_FLAG)
        {
            ILD_DUTY.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ILD_DUTY.SetLookupParamValue("W_WHERE", P_WHERE);
            ILD_DUTY.SetLookupParamValue("W_ENABLED_FLAG_YN", P_ENABLED_FLAG);
        }

        #endregion


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

        #endregion;

        #region ----- Initialize -----

        private void Insert_Before_Apply_Day()
        {
        }

        private bool Insert_Holy_Type()
        {
            if (WORK_DATE_FR.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_FR.Focus();
                return false;
            }
            if (WORK_DATE_TO.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_TO.Focus();
                return false;
            }
            if (WORK_TYPE_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Type(교대 유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_TYPE.Focus();
                return false;
            }
            return true;
        }                
        
        private void Init_Cell_Status()
        {
            if (iConv.ISNull(IGR_CALENDAR_SET.GetCellValue("HOLY_TYPE")) == iConv.ISNull("-1"))
            {
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE")].Insertable = 0;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE")].Updatable = 0;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE_NAME")].Insertable = 0;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE_NAME")].Updatable = 0;

                IGR_CALENDAR_SET.CurrentCellMoveTo(3);
            }
            else
            {
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE")].Insertable = 1;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE")].Updatable = 1;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE_NAME")].Insertable = 1;
                IGR_CALENDAR_SET.GridAdvExColElement[IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE_NAME")].Updatable = 1;

                IGR_CALENDAR_SET.CurrentCellMoveTo(1);
            }
        }

        private void Init_Work_Plan()
        {
            //기존 자료 삭제.
            IGR_CALENDAR_SET.BeginUpdate();
            IDA_CALENDAR_SET.OraSelectData.AcceptChanges();
            for (int i = 0; i < IDA_CALENDAR_SET.SelectRows.Count; i++)
            {
                IDA_CALENDAR_SET.OraSelectData.Rows[i].SetAdded();
            }
            IDA_CALENDAR_SET.Refillable = false;
            IGR_CALENDAR_SET.EndUpdate();
            
            IGR_CALENDAR_SET.CurrentCellMoveTo(0, 0);
            IGR_CALENDAR_SET.CurrentCellActivate(0, 0);
            IGR_CALENDAR_SET.Focus();
        }

        private void Init_Work_Plan_STD()
        {
            IDA_WORK_PLAN_STD.Fill();
            if (IDA_WORK_PLAN_STD.SelectRows.Count == 0)
            {
                return;
            }

            IGR_CALENDAR_SET.BeginUpdate();
            for (int i = 0; i < IDA_WORK_PLAN_STD.SelectRows.Count; i++)
            {
                IDA_CALENDAR_SET.AddUnder();
                for (int j = 0; j < IGR_CALENDAR_SET.GridAdvExColElement.Count; j++)
                {
                    IGR_CALENDAR_SET.SetCellValue(i, j, IDA_WORK_PLAN_STD.SelectRows[i][j]);
                }
            }
            IGR_CALENDAR_SET.EndUpdate();

            IGR_CALENDAR_SET.CurrentCellMoveTo(0, 0);
            IGR_CALENDAR_SET.CurrentCellActivate(0, 0);
            IGR_CALENDAR_SET.Focus();            
        }

        private Boolean Save_Work_Plan()
        {
            string mCREATED_METHOD;
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 --             
            //if (iString.ISNull(PERSON_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "P".ToString();
            //}
            //else if (iString.ISNull(DEPT_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "D".ToString();
            //}
            //else
            //{
            //    mCREATED_METHOD = "A".ToString();
            //}
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 -- 
            mCREATED_METHOD = "A".ToString();

            try
            {
                // 기존자료 삭제.
                IDC_DELETE_CALENDAR_SET_ALL.SetCommandParamValue("W_CREATED_METHOD", mCREATED_METHOD);
                IDC_DELETE_CALENDAR_SET_ALL.ExecuteNonQuery();

                Init_Work_Plan();

                //기적용일수 저장.
                IDC_SAVE_PRE_WORK_DAY.SetCommandParamValue("P_CREATED_METHOD", mCREATED_METHOD);
                IDC_SAVE_PRE_WORK_DAY.ExecuteNonQuery();

                //근무 일정 저장.
                IDA_CALENDAR_SET.SetInsertParamValue("P_CREATED_METHOD", mCREATED_METHOD);
                IDA_CALENDAR_SET.SetDeleteParamValue("W_CREATED_METHOD", mCREATED_METHOD);
                IDA_CALENDAR_SET.Update();
            }
            catch (Exception EX)
            {
                MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }


        #region ----- Excel Upload -----

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog1.DefaultExt = "xlsx";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private bool Excel_Upload()
        {
            bool vResult = false;

            if (iConv.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(V_START_ROW.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_START_ROW))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            if (iConv.ISNull(WORK_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WORK_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool vXL_Load_OK = false;
            string vOPenFileName = UPLOAD_FILE_PATH.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;
            
            V_MESSAGE.PromptText = "Importing Start....";
            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL_Detail(IDC_IMPORT_CALENDAR_DETAIL, iConv.ISNumtoZero(V_START_ROW.EditValue, 2), V_PB_INTERFACE, V_MESSAGE);
                    if (vXL_Load_OK == false)
                    {
                        vResult = false;
                    }
                    else
                    {
                        V_MESSAGE.PromptText = "Importing Completed....";
                        vResult = true;
                    }
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                vResult = false;
                return vResult;
            }
            vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return vResult;
        }

        #endregion

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----

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

        private void HRMF0303_CREATE_Load(object sender, EventArgs e)
        {
            Show_Import(false);
            IDA_CALENDAR_SET.FillSchema();

            //igbCORP_GROUP.BringToFront();
            igbCORP_GROUP.Visible = false; //.Show(); 
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if (iConv.ISNull(G_CORP_TYPE.EditValue) == "ALL")
            {
                igbCORP_GROUP.Visible = true; //.Show();

                irb_ALL.RadioButtonValue = "A";
                G_CORP_TYPE.EditValue = "A";
            }

        }

        private void HRMF0303_CREATE_Shown(object sender, EventArgs e)
        {
            BTN_IMPORT.BringToFront();
            BTN_DAILY_SELECT.BringToFront();
            BTN_DAILY_CANCEL.BringToFront();
            BTN_DAILY_SAVE.BringToFront();

            idcYYYYMM_TERM.SetCommandParamValue("W_YYYYMM", WORK_YYYYMM.EditValue);
            idcYYYYMM_TERM.ExecuteNonQuery();
            WORK_DATE_FR.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_START_DATE");
            WORK_DATE_TO.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_END_DATE");

            PM_NOTIFY.PromptTextElement[0].Default = string.Format("※{0}", isMessageAdapter1.ReturnText("HRM_10015"));
            PM_NOTIFY.Refresh();
        }

        private void igrCALENDAR_SET_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == IGR_CALENDAR_SET.GetColumnToIndex("HOLY_TYPE"))
            {
                IGR_CALENDAR_SET.CurrentCellActivate(IGR_CALENDAR_SET.GetColumnToIndex("DAY_COUNT"));
            }
        }

        private void START_DATE_EditValueChanged(object pSender)
        {
            WORK_DATE_TO.EditValue = iDate.ISMonth_Last(WORK_DATE_FR.EditValue);
        }

        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Save_Work_Plan() == false)
            {
                return;
            }

            Search_DB("A");
        }

        private void ibtCREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;            
            Application.DoEvents();
            
            //월 근무기준 변경 내역 저장 
            if (Save_Work_Plan() == false)
            {
                return;
            }

            string mCREATED_METHOD;
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 --             
            //if (iString.ISNull(PERSON_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "P".ToString();
            //}
            //else if (iString.ISNull(DEPT_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "D".ToString();
            //}
            //else
            //{
            //    mCREATED_METHOD = "A".ToString();
            //}
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 -- 
            mCREATED_METHOD = "A".ToString();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_SET_CALENDAR_DETAIL.SetCommandParamValue("P_CREATE_TYPE", mCREATED_METHOD);
            IDC_SET_CALENDAR_DETAIL.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_CALENDAR_DETAIL.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_CALENDAR_DETAIL.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_CALENDAR_DETAIL.ExcuteError || vSTATUS == "F")
            {
                UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomatioin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            //일근무계획 생성 
            SEARCH_DB_DETAIL();
        }

        private void BTN_CLOSEDL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        private void btnINSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Insert_Holy_Type() == false)
            {
                return;
            }
            IDA_CALENDAR_SET.AddUnder();
            if (IGR_CALENDAR_SET.RowIndex == Convert.ToInt32(0))
            {
                Insert_Before_Apply_Day();
            }
            Init_Cell_Status();
            if (iConv.ISNull(WORK_TYPE_GROUP.EditValue) == "11" 
               && iConv.ISNull(IGR_CALENDAR_SET.GetCellValue("HOLY_TYPE")) != iConv.ISNull("-1"))
            {
                IDA_CALENDAR_SET.Delete();
            }            
        }

        private void btnDELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CALENDAR_SET.Delete();
        }

        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CALENDAR_SET.Cancel();
        }

        private void BTN_DAILY_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SEARCH_DB_DETAIL();
        }

        private void BTN_DAILY_PRE_ButtonClick(object pSender, EventArgs pEventArgs)
        {

        }

        private void BTN_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Import(true);
        }

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }
        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Excel_Upload() == true)
            {
                Show_Import(false);
                SEARCH_DB_DETAIL();
            }
        }

        private void BTN_CLOSED_I_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Import(false); 
        }

        private void BTN_DAILY_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CALENDAR_DETAIL_1.Update();
            IDA_CALENDAR_DETAIL_2.Update();
        }

        private void BTN_DAILY_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_CALENDAR_DETAIL_1.Cancel();
            IDA_CALENDAR_DETAIL_2.Cancel();
        }

        private void BTN_SET_WORKCALENDAR_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(WORK_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_YYYYMM.Focus();
                return;
            }
            if (iConv.ISNull(WORK_TYPE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Type(교대 유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_TYPE.Focus();
                return;
            }
            if (WORK_DATE_FR.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_FR.Focus();
                return;
            }
            if (WORK_DATE_TO.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_TO.Focus();
                return;
            }

            try
            {
                //일 근무기준 update//
                IDA_CALENDAR_DETAIL_1.Update();
                IDA_CALENDAR_DETAIL_2.Update();
            }
            catch(Exception Ex)
            {
                MessageBoxAdv.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (IDA_CALENDAR_DETAIL_1.CurrentRows.Count == 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10420"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string mCREATED_METHOD;
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 --             
            //if (iString.ISNull(PERSON_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "P".ToString();
            //}
            //else if (iString.ISNull(DEPT_ID.EditValue) != string.Empty)
            //{
            //    mCREATED_METHOD = "D".ToString();
            //}
            //else
            //{
            //    mCREATED_METHOD = "A".ToString();
            //}
            //2015.03.30 J.LAKE 변경 -> 근무계획 생성 방법 변경 -- 
            mCREATED_METHOD = "A".ToString();
            
            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_SET_WORKCALENDAR.SetCommandParamValue("P_CREATE_TYPE", mCREATED_METHOD);
            IDC_SET_WORKCALENDAR.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_WORKCALENDAR.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_WORKCALENDAR.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_WORKCALENDAR.ExcuteError || vSTATUS == "F")
            {
                UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vMESSAGE != string.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomatioin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaWORK_TYPE_SelectedRowData(object pSender)
        {
            Search_DB("A");
        }

        private void ilaPERSON_SelectedRowData(object pSender)
        {
            //Search_DB("P");
        }
        
        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_CORP_ID", CORP_ID.EditValue);
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_CORP_ID", CORP_ID.EditValue);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaWORK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("WORK_TYPE", "Y");
        }

        private void ilaHOLY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("HOLY_TYPE", "Y");
        }

        private void ilaHOLY_TYPE2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("HOLY_TYPE", "Y");
        }

        private void ilaHOLY_TYPE3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("HOLY_TYPE", "Y");
        }

        private void ILA_DUTY_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_DUTY_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DUTY.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        
        #endregion

        #region ----- Adapter Event -----

        private void idaCALENDAR_SET_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(WORK_DATE_FR.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(WORK_DATE_TO.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(WORK_TYPE_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Type(교대 유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(WORK_TYPE.EditValue).Substring(0, 2) == "11")
            {
            }
            else
            {

                if (iConv.ISNull(e.Row["HOLY_TYPE"]) == String.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy Type(근무구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (iConv.ISNull(e.Row["HOLY_TYPE_NAME"]) == String.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy Type Name(근무명) "), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (iConv.ISNull(e.Row["DAY_COUNT"]) == String.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Day Count(근무일수) "), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void IDA_CALENDAR_DETAIL_1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["WORK_DATE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DUTY_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Type(근태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["HOLY_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy Type(근무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void IDA_CALENDAR_DETAIL_2_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["WORK_DATE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["DUTY_ID"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Duty Type(근태)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["HOLY_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Holy Type(근무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void idaCALENDAR_SET_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            Init_Cell_Status();
        }


        #endregion

        private void irb_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

    }
}