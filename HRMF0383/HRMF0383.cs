using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0383
{
    public partial class HRMF0383 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        private ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        private ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #endregion;

        #region ----- Constructor -----

        public HRMF0383()
        {
            InitializeComponent();
        }

        public HRMF0383(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if(iString.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE_0.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
                CORP_TYPE_1.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
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

        #endregion;

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {                    
                    if (TB_MAIN.SelectedTab.TabIndex == TP_WORK_TYPE.TabIndex)
                    {
                        SearchWorkType();
                    }
                    else if (TB_MAIN.SelectedTab.TabIndex == TP_FLOOR.TabIndex)
                    {
                        SearchFloor(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_CHANGE_WORK_TYPE.IsFocused)
                    {
                        IDA_CHANGE_WORK_TYPE.Update();
                    }
                    else if (IDA_CHANGE_HISTORY_WT.IsFocused)
                    {
                        IDA_CHANGE_HISTORY_WT.Update();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                    if (IDA_CHANGE_WORK_TYPE.IsFocused)
                    {
                        IDA_CHANGE_WORK_TYPE.Cancel();
                    }
                    else if (IDA_CHANGE_HISTORY_WT.IsFocused)
                    {
                        IDA_CHANGE_HISTORY_WT.Cancel();
                    }
                    else if (IDA_WORK_CALENDAR.IsFocused)
                    {
                        IDA_WORK_CALENDAR.Cancel();
                    }
                    else if (idaMODIFY_FLOOR.IsFocused)
                    {
                        idaMODIFY_FLOOR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CHANGE_HISTORY_WT.IsFocused)
                    {
                        IDA_CHANGE_HISTORY_WT.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                }
            }
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

        #region ----- Private Method ----

        private void DefaultCorporation()
        {
            W_CHANGE_DATE.EditValue = System.DateTime.Today;
            STD_DATE_1.EditValue = System.DateTime.Today;

            //System.DateTime vDate = new System.DateTime(2011, 7, 14);
            //STD_DATE_0.EditValue = vDate;
            //STD_DATE_1.EditValue = vDate;

            // 조회년월 SETTING
            ildYYYYMM_0.SetLookupParamValue("W_START_YYYYMM", "2010-01");

            //WORK_YYYYMM_2.EditValue = ISDate.ISYearMonth(DateTime.Today);
            WORK_YYYYMM_2.EditValue = iDate.ISYearMonth(W_CHANGE_DATE.DateTimeValue);
            idcYYYYMM_TERM.SetCommandParamValue("W_YYYYMM", WORK_YYYYMM_2.EditValue);
            idcYYYYMM_TERM.ExecuteNonQuery();
            DATE_SEARCH_START_1.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_START_DATE");
            DATE_CREATE_START_2.EditValue = W_CHANGE_DATE.DateTimeValue;
            DATE_CREATE_END_2.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_END_DATE");

            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // CORP_TYPE이면 그룹박스표시 

            CORP_NAME_1 .BringToFront();
            W_CORP_NAME_1.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            igbCORP_GROUP_1.BringToFront();
            igbCORP_GROUP_0.Visible = false; //.Show();
            igbCORP_GROUP_1.Visible = false;

            if (iString.ISNull(CORP_TYPE_0.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();
                igbCORP_GROUP_1.Visible = true;

                irb_ALL_0.RadioButtonValue = "A";
                irb_ALL_1.RadioButtonValue = "A";
                CORP_TYPE_0.EditValue = "A";
                CORP_TYPE_1.EditValue = "A";

            }
            else if (iString.ISNull(CORP_TYPE_0.EditValue) == "1")
            {
                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();
                W_CORP_NAME_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
                CORP_NAME_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_1.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                CORP_NAME_1.EditValue = W_CORP_NAME_1.EditValue;
                CORP_ID_1.EditValue = W_CORP_ID_1.EditValue;
            }

  


            //작업장
            idcDEFAULT_FLOOR.ExecuteNonQuery();
            W_FLOOR_NAME_1.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            W_FLOOR_ID_1.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");
            FLOOR_NAME_1.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            FLOOR_ID_1.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");

            object oPERSON_NAME = idcDEFAULT_FLOOR.GetCommandParamValue("O_PERSON_NAME");
            object oCAPACITY = idcDEFAULT_FLOOR.GetCommandParamValue("O_CAPACITY"); //권한
            string vCAPACITY = ConvertString(oCAPACITY);





            //FLOOR_NAME_0.EditValue = "후가공";
            //FLOOR_ID_0.EditValue = 3707;

            //FLOOR_NAME_1.EditValue = "후가공";
            //FLOOR_ID_1.EditValue = 3707;


            
            ////인사담당자이면 -- 담당자의 담당하는 작업장만 보게 하려고
            //if (vCAPACITY == "C")
            //{
            //    FLOOR_NAME_0.ReadOnly = false;
            //    isGroupBox1.PromptTextElement[0].TL1_KR = string.Format("{0} - {1}[{2}]", isGroupBox1.PromptText, oPERSON_NAME, "인사담당");
            //}
            //else
            //{
            //    FLOOR_NAME_0.ReadOnly = true;
            //    isGroupBox1.PromptTextElement[0].TL1_KR = string.Format("{0} - {1}[{2}]", isGroupBox1.PromptText, oPERSON_NAME, FLOOR_NAME_0.EditValue);
            //}
        }

        private void DefaultEmploye()
        {
            idcDEFAULT_EMPLOYE_TYPE_0.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            idcDEFAULT_EMPLOYE_TYPE_0.ExecuteNonQuery();
            W_EMPLOYE_TYPE_NAME_1.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE_NAME");
            W_EMPLOYE_TYPE_1.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE");

            EMPLOYE_TYPE_NAME_1.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE_NAME");
            EMPLOYE_TYPE_1.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE");
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SearchFloor()
        {
            ISCommonUtil.ISFunction.ISConvert vString = new ISCommonUtil.ISFunction.ISConvert();

            object vObject1 = FLOOR_NAME_1.EditValue;
            object vObject2 = PERSON_NAME_1.EditValue;
            if (vString.ISNull(vObject1) == string.Empty && vString.ISNull(vObject2) == string.Empty)
            {
                //검색 조건을 선택 하세요!
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10305"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                idaMODIFY_FLOOR.Fill();
            }
            catch(System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void SearchWorkType()
        {
            if (TB_WORK_TYPE.SelectedTab.TabIndex == TP_WORK_CALENDAR.TabIndex)
            {
                SearchWorkcalendar();
            }
            else
            {
                if (iConv.ISNull(W_FLOOR_NAME_1.EditValue) == string.Empty &&
                    iConv.ISNull(W_WORK_TYPE_ID_1.EditValue) == string.Empty &&
                    iConv.ISNull(W_PERSON_ID_1.EditValue) == string.Empty)
                {
                    //검색 조건을 선택 하세요!
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10305"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                IDA_CHANGE_WORK_TYPE.Fill();
                IGR_CHANGE_WORK_TYPE.Focus();
            }
        }

        private void SearchWorkcalendar()
        {
            if (iConv.ISNull(V_PERSON_ID.EditValue) == string.Empty)
            {
                //검색 조건을 선택 하세요!
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10305"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_WORK_CALENDAR.Fill();
            IGR_WORK_CALENDAR.Focus();
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0383_Load(object sender, EventArgs e)
        {
            DefaultCorporation();
            DefaultEmploye();
            GB_CHANGE_INFO.BringToFront();

            IDA_CHANGE_WORK_TYPE.FillSchema();
        }

        private void BTN_EXCEL_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0383_UPLOAD vHRMF0383_UPLOAD = new HRMF0383_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface, W_CORP_ID_1.EditValue);
            vdlgResult = vHRMF0383_UPLOAD.ShowDialog();
            vHRMF0383_UPLOAD.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                SearchWorkType();
            }
        }

        #endregion;

        #region ----- Grid Event ----

        private void IGR_CHANGE_WORK_TYPE_CellDoubleClick(object pSender)
        {
            int vTabIndex = TP_WORK_CALENDAR.TabIndex;

            V_PERSON_NAME.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("NAME");
            V_PERSON_ID.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("PERSON_ID");
            V_PERSON_NUM.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("PERSON_NUM");
            V_JOB_CATEGORY_NAME.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("JOB_CATEGORY_NAME");
            V_JOIN_DATE.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("JOIN_DATE");
            V_RETIRE_DATE.EditValue = IGR_CHANGE_WORK_TYPE.GetCellValue("RETIRE_DATE");

            object vCHANGE_DATE = IGR_CHANGE_WORK_TYPE.GetCellValue("CHANGE_DATE");
            V_START_DATE.EditValue = iDate.ISMonth_1st(vCHANGE_DATE);
            V_END_DATE.EditValue = iDate.ISMonth_Last(V_START_DATE.EditValue);

            TB_WORK_TYPE.SelectedIndex = (vTabIndex - 1);
            SearchWorkcalendar(); 
        }

        #endregion;

        #region ----- Method Event -----
         
        private void MODIFY_FLOOR_Save()
        {
            string vMessage = string.Empty;
            idaMODIFY_FLOOR.Update();

            object vObject = O_SUCCESS_FLAG_2.EditValue;
            string vSuccess = ConvertString(vObject);

            if (vSuccess == "Y")
            {

                try
                {
                    //FCM_10344 //작업장을 수정 하였습니다.
                    //FCM_10341 //다시 조회를 하십시오!
                    vMessage = string.Format("{0}\n\n{1}", isMessageAdapter1.ReturnText("FCM_10344"), isMessageAdapter1.ReturnText("FCM_10341"));
                    MessageBoxAdv.Show(vMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBoxAdv.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                try
                {
                    vMessage = string.Format("{0}", O_MESSAGE_2.EditValue);
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (System.Exception ex)
                {
                    MessageBoxAdv.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        //private void Delete_WORK_TYPE()
        //{
            //string vMessage = string.Empty;

            //object vObject = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("LAST_YN");
            //string vLastYN = ConvertString(vObject);

            //int vCountRow = igrMODIFY_HISTORY_WORKTYPE.RowCount;

            //if (vLastYN == "Y" && vCountRow > 1)
            //{
            //    System.Windows.Forms.DialogResult vChoice;

            //    object vObject_PERSON_NAME = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("PERSON_NAME");
            //    object vObject_PERSON_NUMBER = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("PERSON_NUMBER");
            //    object vObject_WORK_TYPE_NAME = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("H_WORK_TYPE_NAME");
            //    object vObject_EFFECTIVE_DATE_FR = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("EFFECTIVE_DATE_FR");

            //    string vPERSON_NAME = ConvertString(vObject_PERSON_NAME);
            //    string vPERSON_NUMBER = ConvertString(vObject_PERSON_NUMBER);
            //    string vWORK_TYPE_NAME = ConvertString(vObject_WORK_TYPE_NAME);
            //    System.DateTime vEFFECTIVE_DATE_FR = ConvertDateTime(vObject_EFFECTIVE_DATE_FR);

            //    //삭제 하시겠습니까?
            //    vMessage = string.Format("{0}[{1}]\n{2}\n{3}\n\n{4}", vPERSON_NAME, vPERSON_NUMBER, vWORK_TYPE_NAME, vEFFECTIVE_DATE_FR.ToShortDateString(), isMessageAdapter1.ReturnText("EAPP_10030"));

            //    vChoice = MessageBoxAdv.Show(vMessage, "Delete", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button2);

            //    if (vChoice == System.Windows.Forms.DialogResult.Yes)
            //    {
            //        try
            //        {
            //            object vW_WORK_CORP_ID = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("WORK_CORP_ID");
            //            object vW_PERSON_ID = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("PERSON_ID");
            //            object vW_EFFECTIVE_DATE_FR = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("EFFECTIVE_DATE_FR");
            //            object vW_EFFECTIVE_DATE_TO = igrMODIFY_HISTORY_WORKTYPE.GetCellValue("EFFECTIVE_DATE_TO");

            //            idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_MODIFY_TAB", "W");
            //            idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_WORK_CORP_ID", vW_WORK_CORP_ID);
            //            idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_PERSON_ID", vW_PERSON_ID);
            //            idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_EFFECTIVE_DATE_FR", vW_EFFECTIVE_DATE_FR);
            //            idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_EFFECTIVE_DATE_TO", vW_EFFECTIVE_DATE_TO);

            //            idcDELETE_PERSON_HISTORY.ExecuteNonQuery();

            //            object vObject_DELETE_SUCCESS_FLAG = idcDELETE_PERSON_HISTORY.GetCommandParamValue("O_DELETE_SUCCESS_FLAG");
            //            object vObject_MODIFY_SUCCESS_WORK_TYPE = idcDELETE_PERSON_HISTORY.GetCommandParamValue("O_MODIFY_SUCCESS_WORK_TYPE");

            //            string vDELETE_SUCCESS_FLAG = ConvertString(vObject_DELETE_SUCCESS_FLAG);
            //            string vMODIFY_SUCCESS_WORK_TYPE = ConvertString(vObject_MODIFY_SUCCESS_WORK_TYPE);

            //            if (vDELETE_SUCCESS_FLAG == "Y")
            //            {
            //                //삭제 하였습니다. [ FCM_10356]
            //                //다시 조회를 하십시오! [ FCM_10341]
            //                vMessage = string.Format("{0}\n\n{1}\n\n{2}", isMessageAdapter1.ReturnText("FCM_10356"), vMODIFY_SUCCESS_WORK_TYPE, isMessageAdapter1.ReturnText("FCM_10341"));
            //                MessageBoxAdv.Show(vMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            }
            //        }
            //        catch (System.Exception ex)
            //        {
            //            isAppInterfaceAdv1.OnAppMessage(ex.Message);
            //            System.Windows.Forms.Application.DoEvents();
            //        }
            //    }
            //}
            //else
            //{
            //    //삭제할 수 없습니다! [EAPP_10013]
            //    vMessage = string.Format("{0}", isMessageAdapter1.ReturnText("EAPP_10013"));
            //    MessageBoxAdv.Show(vMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        //}

        //private void Delete_FLOOR()
        //{
        //    string vMessage = string.Empty;

        //    object vObject = igrMODIFY_HISTORY_FLOOR.GetCellValue("LAST_YN");
        //    string vLastYN = ConvertString(vObject);

        //    int vCountRow = igrMODIFY_HISTORY_FLOOR.RowCount;

        //    if (vLastYN == "Y" && vCountRow > 1)
        //    {
        //        System.Windows.Forms.DialogResult vChoice;

        //        object vObject_PERSON_NAME = igrMODIFY_HISTORY_FLOOR.GetCellValue("PERSON_NAME");
        //        object vObject_PERSON_NUMBER = igrMODIFY_HISTORY_FLOOR.GetCellValue("PERSON_NUMBER");
        //        object vObject_FLOOR_NAME = igrMODIFY_HISTORY_FLOOR.GetCellValue("H_FLOOR_NAME");
        //        object vObject_EFFECTIVE_DATE_FR = igrMODIFY_HISTORY_FLOOR.GetCellValue("EFFECTIVE_DATE_FR");

        //        string vPERSON_NAME = ConvertString(vObject_PERSON_NAME);
        //        string vPERSON_NUMBER = ConvertString(vObject_PERSON_NUMBER);
        //        string vFLOOR_NAME = ConvertString(vObject_FLOOR_NAME);
        //        System.DateTime vEFFECTIVE_DATE_FR = ConvertDateTime(vObject_EFFECTIVE_DATE_FR);

        //        //삭제 하시겠습니까?
        //        vMessage = string.Format("{0}[{1}]\n{2}\n{3}\n\n{4}", vPERSON_NAME, vPERSON_NUMBER, vFLOOR_NAME, vEFFECTIVE_DATE_FR.ToShortDateString(), isMessageAdapter1.ReturnText("EAPP_10030"));
        //        vChoice = MessageBoxAdv.Show(vMessage, "Delete", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button2);

        //        if (vChoice == System.Windows.Forms.DialogResult.Yes)
        //        {
        //            try
        //            {
        //                object vW_WORK_CORP_ID = igrMODIFY_HISTORY_FLOOR.GetCellValue("WORK_CORP_ID");
        //                object vW_PERSON_ID = igrMODIFY_HISTORY_FLOOR.GetCellValue("PERSON_ID");
        //                object vW_EFFECTIVE_DATE_FR = igrMODIFY_HISTORY_FLOOR.GetCellValue("EFFECTIVE_DATE_FR");
        //                object vW_EFFECTIVE_DATE_TO = igrMODIFY_HISTORY_FLOOR.GetCellValue("EFFECTIVE_DATE_TO");

        //                idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_MODIFY_TAB", "F");
        //                idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_WORK_CORP_ID", vW_WORK_CORP_ID);
        //                idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_PERSON_ID", vW_PERSON_ID);
        //                idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_EFFECTIVE_DATE_FR", vW_EFFECTIVE_DATE_FR);
        //                idcDELETE_PERSON_HISTORY.SetCommandParamValue("W_EFFECTIVE_DATE_TO", vW_EFFECTIVE_DATE_TO);

        //                idcDELETE_PERSON_HISTORY.ExecuteNonQuery();

        //                object vObject_DELETE_SUCCESS_FLAG = idcDELETE_PERSON_HISTORY.GetCommandParamValue("O_DELETE_SUCCESS_FLAG");

        //                string vDELETE_SUCCESS_FLAG = ConvertString(vObject_DELETE_SUCCESS_FLAG);

        //                if (vDELETE_SUCCESS_FLAG == "Y")
        //                {
        //                    //삭제 하였습니다. [ FCM_10356]
        //                    //다시 조회를 하십시오! [ FCM_10341]
        //                    vMessage = string.Format("{0}\n\n{1}", isMessageAdapter1.ReturnText("FCM_10356"), isMessageAdapter1.ReturnText("FCM_10341"));
        //                    MessageBoxAdv.Show(vMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                }
        //            }
        //            catch (System.Exception ex)
        //            {
        //                isAppInterfaceAdv1.OnAppMessage(ex.Message);
        //                System.Windows.Forms.Application.DoEvents();
        //            }
        //        }
        //    }
        //    else
        //    {
        //        //삭제할 수 없습니다! [EAPP_10013]
        //        vMessage = string.Format("{0}", isMessageAdapter1.ReturnText("EAPP_10013"));
        //        MessageBoxAdv.Show(vMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}

        #endregion;

        #region ----- LookUP Event ----

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_W_0.SetLookupParamValue("W_START_DATE", W_CHANGE_DATE.EditValue);
            ildPERSON_W_0.SetLookupParamValue("W_END_DATE", W_CHANGE_DATE.EditValue);
        }

        private void ilaWORK_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaWORK_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaEMPLOYE_TYPE_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaYYYYMM_0_SelectedRowData(object pSender)
        {
            object vObject = WORK_YYYYMM_2.EditValue;
            string vYYYYMM = ConvertString(vObject);
            if (string.IsNullOrEmpty(vYYYYMM) == false)
            {
                System.DateTime v1stDate = iDate.ISMonth_1st(vYYYYMM);
                System.DateTime vLastDate = iDate.ISMonth_Last(vYYYYMM);
                DATE_CREATE_START_2.EditValue = v1stDate.ToShortDateString();
                DATE_CREATE_END_2.EditValue = vLastDate.ToShortDateString();
            }
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

        private void ilaOCPT_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOCPT_2.SetLookupParamValue("W_GROUP_CODE", "OCPT");
            ildOCPT_2.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaHOLY_TYPE_SelectedRowData(object pSender)
        {
            object mOPEN_TIME;
            object mCLOSE_TIME;
            idcWORK_IO_TIME.ExecuteNonQuery();
            mOPEN_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_OPEN_TIME");
            mCLOSE_TIME = idcWORK_IO_TIME.GetCommandParamValue("O_CLOSE_TIME");
            igrWORK_CALENDAR.SetCellValue("OPEN_TIME", mOPEN_TIME);
            igrWORK_CALENDAR.SetCellValue("CLOSE_TIME", mCLOSE_TIME);
        }

        private void STD_DATE_EditValueChanged(object pSender)
        {
            string vYYYYMM = string.Format("{0:D4}-{1:D2}", W_CHANGE_DATE.DateTimeValue.Year, W_CHANGE_DATE.DateTimeValue.Month);
            System.DateTime v1stDate = iDate.ISMonth_1st(vYYYYMM);
            System.DateTime vLastDate = iDate.ISMonth_Last(vYYYYMM);
            WORK_YYYYMM_2.EditValue = vYYYYMM;

            DATE_SEARCH_START_1.EditValue = v1stDate.ToShortDateString();

            DATE_CREATE_START_2.DateTimeValue = W_CHANGE_DATE.DateTimeValue;
            DATE_CREATE_END_2.EditValue = vLastDate.ToShortDateString();
        }

        #endregion

        #region ----- Adapter Event ----

        private void idaMODIFY_FLOOR_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            IDA_CHANGE_HISTORY_WT.Fill();
        }

        private void IDA_CHANGE_WORK_TYPE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }
            if (iConv.ISNull(e.Row["CHANGE_DATE"]) == string.Empty)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10223"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 
        }

        private void IDA_CHANGE_WORK_TYPE_UpdateCompleted(object pSender)
        {
            //FCM_10296 //근무 계획표 수정 하였습니다.
            //FCM_10341 //다시 조회를 하십시오!            
            MessageBoxAdv.Show(string.Format("{0} \r\n {1}", isMessageAdapter1.ReturnText("FCM_10296"), isMessageAdapter1.ReturnText("FCM_10341")), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void idaMODIFY_WORK_TYPE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            IDA_CHANGE_HISTORY_WT.Fill(); 
        }

        private void ilaCOST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCOST_CENTER_SelectedRowData(object pSender)
        {
            System.Windows.Forms.SendKeys.Send("{TAB}");
        }


        #endregion

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_0.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void irb_ALL_1_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE_1.EditValue = RB_STATUS.RadioCheckedString;
        }

    }
}