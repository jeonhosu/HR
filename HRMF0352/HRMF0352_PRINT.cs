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

namespace HRMF0352
{
    public partial class HRMF0352_PRINT : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string mOUTPUT_TYPE = null;
        #endregion;

        #region ----- Constructor -----

        public HRMF0352_PRINT()
        {
            InitializeComponent();
        }

        public HRMF0352_PRINT(ISAppInterface pAppInterface
                            , string pOUTPUT_TYPE
                            , object pCORP_ID, object pCORP_NAME
                            , object pSTART_DATE, object pEND_DATE
                            , object pDEPT_ID, object pDEPT_NAME
                            , object pFLOOR_ID, object pFLOOR_NAME
                            , object pPERSON_ID, object pPERSON_NUM, object pNAME)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mOUTPUT_TYPE = pOUTPUT_TYPE;

            V_CORP_ID.EditValue = pCORP_ID;
            V_CORP_NAME.EditValue = pCORP_NAME;
            V_START_DATE.EditValue = pSTART_DATE;
            V_END_DATE.EditValue = pEND_DATE;
            V_DEPT_ID.EditValue = pDEPT_ID;
            V_DEPT_NAME.EditValue = pDEPT_NAME;
            V_FLOOR_ID.EditValue = pFLOOR_ID;
            V_FLOOR_NAME.EditValue = pFLOOR_NAME;
            V_PERSON_ID.EditValue = pPERSON_ID;
            V_PERSON_NUM.EditValue = pPERSON_NUM;
            V_NAME.EditValue = pNAME;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(V_CORP_ID.EditValue) == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(V_START_DATE.EditValue) == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_START_DATE.Focus();
                return;
            }
            if (iString.ISNull(V_END_DATE.EditValue) == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                V_END_DATE.Focus();
                return;
            }
            
            CB_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            IGR_SELECT_PERSON.LastConfirmChanges();
            IDA_SELECT_PERSON.OraSelectData.AcceptChanges();
            IDA_SELECT_PERSON.Refillable = true;
                        
            IDA_SELECT_PERSON.Fill();
            IGR_SELECT_PERSON.Focus();

            //// 그리드 부분 업데이트 처리
            //idaMONTH_PAYMENT.OraSelectData.AcceptChanges();
            //idaMONTH_PAYMENT.Refillable = true;

            //idaMONTH_PAYMENT.Fill();
            //idaALLOWANCE_INFO.Fill();
            //idaDEDUCTION_INFO.Fill();
            //idaDUTY_INFO.Fill();
        }

        private void Set_CheckBox()
        {
            int mIDX_Col = IGR_SELECT_PERSON.GetColumnToIndex("SELECT_YN");
            object mCheck_YN = CB_SELECT_YN.CheckBoxValue;
            for (int r = 0; r < IGR_SELECT_PERSON.RowCount; r++)
            {
                IGR_SELECT_PERSON.SetCellValue(r, mIDX_Col, mCheck_YN);
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

        // 인쇄 부분
        // Print 관련 소스 코드 2011.1.15(토)
        // Print 관련 소스 코드 2011.5.11(수) 수정
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

        #region ----- XL Print 1 Method ----

        private void XLPrinting_1(string pOUTPUT_TYPE)
        {
            System.DateTime vStartTime = DateTime.Now;

            object vObject = null;
            string vMessage = string.Empty;
            string vBoxCheck = string.Empty;

            int vSelect_Count = 0;
            int vCountRow = IGR_SELECT_PERSON.RowCount;

            if (vCountRow < 1)
            {
                vMessage = string.Format("Not found printing data");
                isAppInterfaceAdv1.OnAppMessage(vMessage);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //전호수 주석 : 체크박스의 기본 설정값 읽어오기
            //string vCheckedString = igrMONTH_PAYMENT.GridAdvExColElement[vIndexCheckBox].CheckedString;

            int vIDX_SELECT_YN = IGR_SELECT_PERSON.GetColumnToIndex("SELECT_YN");
            //-------------------------------------------------------------------------------------
            for (int vRow = 0; vRow < vCountRow; vRow++)
            {
                vObject = IGR_SELECT_PERSON.GetCellValue(vRow, vIDX_SELECT_YN);
                if (iString.ISNull(vObject) == "Y")
                {
                    vSelect_Count++;
                }
            }

            if (vSelect_Count < 1)
            {
                vMessage = string.Format("Not Selected data");
                isAppInterfaceAdv1.OnAppMessage(vMessage);
                System.Windows.Forms.Application.DoEvents();
                return;
            }
            //-------------------------------------------------------------------------------------
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessage = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessage);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0352_001.xls";
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(pOUTPUT_TYPE, IGR_SELECT_PERSON, IDA_DAY_LEAVE_PERSON);
            }
            catch (System.Exception ex)
            {
                vMessage = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessage);
                System.Windows.Forms.Application.DoEvents();
            }
            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            IGR_SELECT_PERSON.LastConfirmChanges();
            IDA_SELECT_PERSON.OraSelectData.AcceptChanges();
            IDA_SELECT_PERSON.Refillable = true;

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessage = string.Format("Printing End [Total Page : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessage);
            System.Windows.Forms.Application.DoEvents();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_1("PRINT"); // 출력 함수 호출

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                    // 인쇄 완료 메시지 출력
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_1("FILE"); // 출력 함수 호출
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0352_Load(object sender, EventArgs e)
        {
            IDA_SELECT_PERSON.FillSchema();
        }

        private void HRMF0352_PRINT_Shown(object sender, EventArgs e)
        {
            if (iString.ISNull(V_CORP_ID.EditValue) == String.Empty)
            {
                // Lookup SETTING
                ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
                idcDEFAULT_CORP.ExecuteNonQuery();

                V_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                V_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }
            CB_SELECT_YN.CheckedState = ISUtil.Enum.CheckedState.Unchecked;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            isAppInterfaceAdv1.OnAppMessage("");
        }

        private void BTN_FILL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SearchDB();
        }

        private void BTN_PRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            XLPrinting_1(mOUTPUT_TYPE);
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Set_CheckBox();
        }

        private void IGR_SELECT_PERSON_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_CHECK_FLAG = IGR_SELECT_PERSON.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_CHECK_FLAG)
            {
                IGR_SELECT_PERSON.LastConfirmChanges();
                IDA_SELECT_PERSON.OraSelectData.AcceptChanges();
                IDA_SELECT_PERSON.Refillable = true;
            }
        }

        #endregion

        #region ----- Lookup Event ----- 

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        #endregion

    }
}