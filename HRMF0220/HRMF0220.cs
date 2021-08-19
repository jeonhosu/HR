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

namespace HRMF0220
{
    public partial class HRMF0220 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0220()
        {
            InitializeComponent();
        }

        public HRMF0220(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        private void DefaultCorporation()
        {
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
                        
            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Insert_Dispatch_Person()
        {
            IGR_PERSON_INFO.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            IGR_PERSON_INFO.SetCellValue("CORP_NAME", CORP_NAME_0.EditValue); 
            IGR_PERSON_INFO.SetCellValue("JOIN_DATE", DateTime.Today);
            IGR_PERSON_INFO.SetCellValue("PAY_DATE", DateTime.Today);

            idcEMPLOYE_STATUS.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            idcEMPLOYE_STATUS.ExecuteNonQuery();
            IGR_PERSON_INFO.SetCellValue("EMPLOYE_TYPE", idcEMPLOYE_STATUS.GetCommandParamValue("O_CODE"));
            IGR_PERSON_INFO.SetCellValue("EMPLOYE_TYPE_NAME", idcEMPLOYE_STATUS.GetCommandParamValue("O_CODE_NAME"));

            IGR_PERSON_INFO.CurrentCellMoveTo(IGR_PERSON_INFO.GetColumnToIndex("PERSON_NUM"));
            IGR_PERSON_INFO.Focus();
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }

        private void Search_DB()
        {
            if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            string vPERSON_ID = iString.ISNull(IGR_PERSON_INFO.GetCellValue("PERSON_ID"));
            int vIDX_PERSO_ID = IGR_PERSON_INFO.GetColumnToIndex("PERSON_ID");
            IDA_PERSON.Fill();
            IGR_PERSON_INFO.Focus();

            for (int i = 0; i < IGR_PERSON_INFO.RowCount; i++)
            {
                if (vPERSON_ID == iString.ISNull(IGR_PERSON_INFO.GetCellValue(i, vIDX_PERSO_ID)))
                {
                    IGR_PERSON_INFO.CurrentCellMoveTo(IGR_PERSON_INFO.GetColumnToIndex("NAME"));
                    return;
                }
            }
        }

        #endregion;

        #region ----- 주민번호 체크 ------
        private bool Repre_Num_Validating_Check(object pRepre_Num)
        {
            if (iString.ISNull(pRepre_Num) == string.Empty)
            {
                return true;
            }
            if (pRepre_Num.ToString().IndexOf("-") == -1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string isReturnValue = null;
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pRepre_Num);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            IGR_PERSON_INFO.SetCellValue("SEX_TYPE", idcREPRE_NUM_CHECK.GetCommandParamValue("O_SEX_TYPE"));
            if (isReturnValue == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            idcSEX_TYPE.SetCommandParamValue("W_GROUP_CODE", "SEX_TYPE");
            idcSEX_TYPE.SetCommandParamValue("W_CODE", IGR_PERSON_INFO.GetCellValue("SEX_TYPE"));
            idcSEX_TYPE.ExecuteNonQuery();
            IGR_PERSON_INFO.SetCellValue("SEX_NAME", idcSEX_TYPE.GetCommandParamValue("O_RETURN_VALUE"));

            if (iString.ISNull(IGR_PERSON_INFO.GetCellValue("BIRTHDAY")) == string.Empty)
            {// 생년월일이 기존에 없을 경우 자동 설정.                
                string mSex_Type = pRepre_Num.ToString().Substring(7, 1);
                if (mSex_Type == "1".ToString() || mSex_Type == "2".ToString() || mSex_Type == "5".ToString() || mSex_Type == "6".ToString())
                {
                    IGR_PERSON_INFO.SetCellValue("BIRTHDAY", DateTime.Parse("19" + pRepre_Num.ToString().Substring(0, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(2, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(4, 2)));
                }
                else
                {
                    IGR_PERSON_INFO.SetCellValue("BIRTHDAY", DateTime.Parse("20" + pRepre_Num.ToString().Substring(0, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(2, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(4, 2)));
                }
                // 음양구분.
                idcCOMMON_W.SetCommandParamValue("W_GROUP_CODE", "BIRTHDAY_TYPE");
                idcCOMMON_W.SetCommandParamValue("W_WHERE", " 1 = 1 ");
                idcCOMMON_W.ExecuteNonQuery();
                IGR_PERSON_INFO.SetCellValue("BIRTHDAY_TYPE_NAME", idcCOMMON_W.GetCommandParamValue("O_CODE_NAME"));
                IGR_PERSON_INFO.SetCellValue("BIRTHDAY_TYPE", idcCOMMON_W.GetCommandParamValue("O_CODE"));
            }
            return true;
        }
        #endregion

        #region ----- Events -----

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
                    if (IDA_PERSON.IsFocused)
                    {
                        IDA_PERSON.AddOver();
                        Insert_Dispatch_Person();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_PERSON.IsFocused)
                    {
                        IDA_PERSON.AddUnder();
                        Insert_Dispatch_Person();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    
                    IDA_PERSON.Update();
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_PERSON.IsFocused)
                    {
                        IDA_PERSON.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_PERSON.IsFocused)
                    {
                        IDA_PERSON.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0220_Load(object sender, EventArgs e)
        {
            
            
        }

        private void HRMF0220_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            CORP_NAME_0.BringToFront();

            //재직구분.
            IDC_DV_COMMON.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            IDC_DV_COMMON.ExecuteNonQuery();
            EMPLOYE_TYPE_0.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE");
            EMPLOYE_TYPE_NAME_0.EditValue = IDC_DV_COMMON.GetCommandParamValue("O_CODE_NAME");

            IDA_PERSON.FillSchema();
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Name"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["WORK_CORP_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Corporation"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["OPERATING_UNIT_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Operating Unit"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["DEPT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NATION_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Nation"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["JOB_CLASS_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Class"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["JOB_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["POST_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Position"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["OCPT_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Ocpt(직무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["ABIL_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Abil(직책)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (string.IsNullOrEmpty(e.Row["REPRE_NUM"].ToString()))
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Repre Num(주민번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["SEX_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Sex Type"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["JOIN_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=입사구분"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["ORI_JOIN_DATE"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Ori Join Date(그룹입사일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["JOIN_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Join Date"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (string.IsNullOrEmpty(e.Row["DIR_INDIR_TYPE"].ToString()))
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Dir/InDir Type(직간접 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (iString.ISNull(e.Row["EMPLOYE_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Employe Status"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RETIRE_DATE"]) != string.Empty && iString.ISNull(e.Row["RETIRE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10170"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RETIRE_DATE"]) == string.Empty && iString.ISNull(e.Row["RETIRE_ID"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10171"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["JOB_CATEGORY_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Category"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (iString.ISNull(e.Row["FLOOR_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work center"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }

        private void idaPERSON_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Person Infomation"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        } 

        #endregion

        #region ----- LOOKUP EVENT -----

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
         
        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }

        private void ilaSEX_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("SEX_TYPE", "Y");
        }

        private void ilaNATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("NATION", "Y");
        }

        private void ilaEMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaPOST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("POST", "Y");
        }

        private void ilaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("FLOOR", "Y");
        }

        private void ilaWORK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("WORK_TYPE", "Y");
        }

        private void ilaJOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("JOB_CATEGORY", "Y");
        }

        private void ilaRETIRE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("RETIRE", "Y");
        }

        #endregion

        #region ----- Cell Validating Event -----

        private void IGR_PERSON_INFO_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == IGR_PERSON_INFO.GetColumnToIndex("JOIN_DATE"))
            {
                IGR_PERSON_INFO.SetCellValue("PAY_DATE", e.NewValue);
            }
        }

        private void igrPERSON_INFO_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            
        }

        #endregion

        #region ----- KeyDown Event -----

        private void PERSON_NUM_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                Search_DB();
            }
        }

        private void REPRE_NUM_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                Search_DB();
            }
        }

        private void NAME_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                Search_DB();
            }
        }

        #endregion

    }
}