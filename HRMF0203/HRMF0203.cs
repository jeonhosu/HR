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

namespace HRMF0203
{
    public partial class HRMF0203 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0203(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property / Method -----
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

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iedSTART_DATE_0.EditValue == null)
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedSTART_DATE_0.Focus();
                return;
            }
            if (iedEND_DATE_0.EditValue == null)
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedEND_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(iedSTART_DATE_0.EditValue) > Convert.ToDateTime(iedEND_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedSTART_DATE_0.Focus();
                return;
            }
            idaHISTORY_HEADER.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idaHISTORY_HEADER.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idaHISTORY_HEADER.Fill();
            igrHEADER.Focus();
        }

        private void SEARCH_SUB_DB()
        {

            if (igrHEADER.GetCellValue("HISTORY_HEADER_ID") == null)
            {
                idaHISTORY_LINE.SetSelectParamValue("W_HISTORY_HEADER_ID", 0);
            }
            else
            {
                idaHISTORY_LINE.SetSelectParamValue("W_HISTORY_HEADER_ID", igrHEADER.GetCellValue("HISTORY_HEADER_ID"));
                History_Info();             // 발령 기초 정보 설정.
            }            
            idaHISTORY_LINE.Fill();
        }

        private void History_Info()
        {
            iedHISTORY_NUM.EditValue = igrHEADER.GetCellValue("HISTORY_NUM");
            iedCHARGE_DATE.EditValue = igrHEADER.GetCellValue("CHARGE_DATE");
            iedCHARGE_NAME.EditValue = igrHEADER.GetCellValue("CHARGE_NAME");
            iedDESCRIPTION.EditValue = igrHEADER.GetCellValue("DESCRIPTION");
        }

        private bool iNewcomer_Check()
        {
            ISCommonUtil.ISFunction.ISConvert iConvert = new ISCommonUtil.ISFunction.ISConvert();
            if (iConvert.ISNull(igrHEADER.GetCellValue("NEWCOMER_YN")) == "Y")
            {
                return true;
            }
            return false;
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
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;
        
        #region ----- Application_MainButtonClick -----
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
                    if (idaHISTORY_HEADER.IsFocused)
                    {                        
                        idaHISTORY_HEADER.AddOver();
                        igrHEADER.SetCellValue("CORP_ID", CORP_ID_0.EditValue);

                        igrHEADER.CurrentCellMoveTo(igrHEADER.GetColumnToIndex("CHARGE_DATE"));
                        igrHEADER.Focus();
                    }
                    else if (idaHISTORY_LINE.IsFocused)
                    {
                        idaHISTORY_LINE.AddOver(); 
                        icbPRINT_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;

                        igrLINE.CurrentCellMoveTo(igrLINE.GetColumnToIndex("NAME"));
                        igrLINE.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaHISTORY_HEADER.IsFocused)
                    {
                        idaHISTORY_HEADER.AddUnder();
                        igrHEADER.SetCellValue("CORP_ID", CORP_ID_0.EditValue);

                        igrHEADER.CurrentCellMoveTo(igrHEADER.GetColumnToIndex("CHARGE_DATE"));
                        igrHEADER.Focus();
                    }
                    else if (idaHISTORY_LINE.IsFocused)
                    {
                        idaHISTORY_LINE.AddUnder(); 
                        icbPRINT_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;

                        igrLINE.CurrentCellMoveTo(igrLINE.GetColumnToIndex("NAME")); 
                        igrLINE.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaHISTORY_HEADER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHISTORY_HEADER.IsFocused)
                    {
                        idaHISTORY_LINE.Cancel();
                        idaHISTORY_HEADER.Cancel();
                    }
                    else if (idaHISTORY_LINE.IsFocused)
                    {
                        idaHISTORY_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHISTORY_HEADER.IsFocused)
                    {
                        idaHISTORY_HEADER.Delete();
                    }
                    else if (idaHISTORY_LINE.IsFocused)
                    {
                        idaHISTORY_LINE.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0203_Load(object sender, EventArgs e)
        {
            DateTime pStart_Date = DateTime.Parse(DateTime.Today.Year.ToString() + "-" + DateTime.Today.Month.ToString() + "-01".ToString());
            DateTime pEnd_Date = DateTime.Today;

            idaHISTORY_HEADER.FillSchema();
            idaHISTORY_LINE.FillSchema();

            iedSTART_DATE_0.EditValue = pStart_Date;
            iedEND_DATE_0.EditValue = pEnd_Date;

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }
        #endregion

        #region ------ Adapter Event -----

        private void idaHISTORY_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["CHARGE_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedCHARGE_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CHARGE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedCHARGE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaHISTORY_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                if (igrLINE.RowCount > Convert.ToInt32(0))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data Exists(발령정보가 존재합니다. 해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaHISTORY_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {                        
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedNAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(igrHEADER.GetCellValue("RETIRE_YN")) != string.Empty)
            {
                if (igrHEADER.GetCellValue("RETIRE_YN").ToString() == "Y".ToString())
                {
                    if (e.Row["RETIRE_ID"] == DBNull.Value)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedNAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Retire Reason(퇴직발령시 퇴직사유)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            else
            {
                if (e.Row["RETIRE_ID"] != DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10039"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

            // 발령후 정보
            if (e.Row["OPERATING_UNIT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedOPERATING_UNIT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                e.Cancel = true;
                return;
            }
            if (e.Row["DEPT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedDEPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOB_CLASS_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedJOB_CLASS_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["JOB_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedJOB_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["POST_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedPOST_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["OCPT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedOCPT_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ABIL_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedABIL_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedPAY_GRADE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOB_CATEGORY_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedJOB_CATEGORY_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FLOOR_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(iedFLOOR_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
        }

        private void idaHISTORY_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                // 신규발령 수정 못함.
                if (iNewcomer_Check() == true)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10030"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }

                if (e.Row["HISTORY_LINE_ID"] == DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                    e.Cancel = true;
                    return;
                }
            }
        }
        #endregion

        #region ----- Lookup Parameter -----
        private void isSetCommonLookUpParameter(string P_GROUP_CODE, string P_CODE_NAME, string P_ENABLED_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", P_CODE_NAME);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_ENABLED_YN);
        }
        #endregion

        #region ----- LookUp PopupShow Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_DEPT_LEVEL", DBNull.Value);
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaCHARGE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("CHARGE", null, "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_NAME", DBNull.Value);
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", iedCHARGE_DATE.EditValue);
        }
                
        private void ilaCHARGE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("CHARGE", null, "Y");
        }

        private void ilaRETIRE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RETIRE", null, "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_DEPT_LEVEL", DBNull.Value);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaOPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaJOB_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB_CLASS", null, "Y");
        }

        private void ilaJOB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB", null, "Y");
        }

        private void ilaPOST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("POST", null, "Y");
        }

        private void ilaOCPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("OCPT", null, "Y");
        }

        private void ilaABIL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ABIL", null, "Y");
        }

        private void ilaPAY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("PAY_GRADE", null, "Y");
        }

        private void ilaJOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB_CATEGORY", null, "Y");
        }

        private void ilaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("FLOOR", null, "Y");
        }

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FLOOR_CC.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_COST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCONTRACT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("CONTRACT_TYPE", null, "Y");
        }

        private void ilaDEPT_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_2.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        #endregion


    }
}