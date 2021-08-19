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

namespace HRMF0221
{
    public partial class HRMF0221 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0221()
        {
            InitializeComponent();
        }

        public HRMF0221(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----
        private void DefaultCorporation()
        {
            ILD_CORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // Lookup SETTING
            ILD_DISPATCH_CORP.SetLookupParamValue("W_CORP_TYPE", "4");
            ILD_DISPATCH_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_WORK_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_WORK_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_WORK_CORP_NAME.BringToFront();
        }

        private void Insert_Dispatch_Person()
        {
            IGR_PERSON.SetCellValue("WORK_CORP_ID", W_WORK_CORP_ID.EditValue);
            IGR_PERSON.SetCellValue("WORK_CORP_NAME", W_WORK_CORP_NAME.EditValue);
            IGR_PERSON.SetCellValue("ORI_JOIN_DATE", DateTime.Today);
            IGR_PERSON.SetCellValue("JOIN_DATE", DateTime.Today);

            IDC_EMPLOYE_STATUS.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            IDC_EMPLOYE_STATUS.ExecuteNonQuery();
            IGR_PERSON.SetCellValue("EMPLOYE_TYPE", IDC_EMPLOYE_STATUS.GetCommandParamValue("O_CODE"));
            IGR_PERSON.SetCellValue("EMPLOYE_TYPE_NAME", IDC_EMPLOYE_STATUS.GetCommandParamValue("O_CODE_NAME"));

            IGR_PERSON.CurrentCellMoveTo(IGR_PERSON.GetColumnToIndex("NAME"));
            IGR_PERSON.Focus();
        }

        private void Set_Common_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }

        private void Search_DB()
        {
            IDA_PERSON.Fill();

            IGR_PERSON.CurrentCellMoveTo(IGR_PERSON.GetColumnToIndex("NAME"));
            IGR_PERSON.Focus();
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
            IDC_REPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pRepre_Num);
            IDC_REPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = IDC_REPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            IGR_PERSON.SetCellValue("SEX_TYPE", IDC_REPRE_NUM_CHECK.GetCommandParamValue("O_SEX_TYPE"));
            if (isReturnValue == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            IDC_SEX_TYPE.SetCommandParamValue("W_GROUP_CODE", "SEX_TYPE");
            IDC_SEX_TYPE.SetCommandParamValue("W_CODE", IGR_PERSON.GetCellValue("SEX_TYPE"));
            IDC_SEX_TYPE.ExecuteNonQuery();
            IGR_PERSON.SetCellValue("SEX_NAME", IDC_SEX_TYPE.GetCommandParamValue("O_RETURN_VALUE"));

            if (iString.ISNull(IGR_PERSON.GetCellValue("BIRTHDAY")) == string.Empty)
            {// 생년월일이 기존에 없을 경우 자동 설정.                
                string mSex_Type = pRepre_Num.ToString().Substring(7, 1);
                if (mSex_Type == "1".ToString() || mSex_Type == "2".ToString() || mSex_Type == "5".ToString() || mSex_Type == "6".ToString())
                {
                    IGR_PERSON.SetCellValue("BIRTHDAY", DateTime.Parse("19" + pRepre_Num.ToString().Substring(0, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(2, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(4, 2)));
                }
                else
                {
                    IGR_PERSON.SetCellValue("BIRTHDAY", DateTime.Parse("20" + pRepre_Num.ToString().Substring(0, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(2, 2)
                                                        + "-".ToString()
                                                        + pRepre_Num.ToString().Substring(4, 2)));
                }
                // 음양구분.
                IDC_COMMON_W.SetCommandParamValue("W_GROUP_CODE", "BIRTHDAY_TYPE");
                IDC_COMMON_W.SetCommandParamValue("W_WHERE", " 1 = 1 ");
                IDC_COMMON_W.ExecuteNonQuery();
                IGR_PERSON.SetCellValue("BIRTHDAY_TYPE_NAME", IDC_COMMON_W.GetCommandParamValue("O_CODE_NAME"));
                IGR_PERSON.SetCellValue("BIRTHDAY_TYPE", IDC_COMMON_W.GetCommandParamValue("O_CODE"));
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

        private void HRMF0221_Load(object sender, EventArgs e)
        {
            DefaultCorporation();

            IDA_PERSON.FillSchema();
        }

        private void HRMF0221_Shown(object sender, EventArgs e)
        {
            //재직구분
            IDC_EMPLOYE_STATUS.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            IDC_EMPLOYE_STATUS.ExecuteNonQuery();
            W_EMPLOYE_TYPE.EditValue = IDC_EMPLOYE_STATUS.GetCommandParamValue("O_CODE");
            W_EMPLOYE_TYPE_NAME.EditValue = IDC_EMPLOYE_STATUS.GetCommandParamValue("O_CODE_NAME");
        }

        #endregion

        #region ----- Adapter Event -----
        private void idaPERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Row["NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Name(성명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(소속 업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Dispatch Corporation(근무 업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["OPERATING_UNIT_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Operating Unit(사업장)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["DEPT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department(부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["NATION_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=국가"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["JOB_CLASS_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Class(직군)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["JOB_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job(직종)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["POST_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Post(직위)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["REPRE_NUM"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Repre Num(주민번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["SEX_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Sex Type(성별)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOIN_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=입사구분"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ORI_JOIN_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Ori Join Date(그룹입사일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOIN_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Join Date(입사일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["DIR_INDIR_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Dir/InDir Type(직간접 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["EMPLOYE_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Employee Status(재직구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            if (e.Row["JOB_CATEGORY_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Category(직구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FLOOR_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Floor(작업장)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPERSON_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Person Infomation(인사정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        } 
        #endregion

        #region ----- LOOKUP EVENT -----

        private void ILA_CORP_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CORP.SetLookupParamValue("W_CORP_TYPE", "1");
            ILD_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ILA_CORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CORP.SetLookupParamValue("W_CORP_TYPE", "1");
            ILD_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_OPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_DISPATCH_CORP_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DISPATCH_CORP.SetLookupParamValue("W_ENABLED_FLAG", "N");
            ILD_DISPATCH_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
        }

        private void ILA_DISPATCH_CORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DISPATCH_CORP.SetLookupParamValue("W_ENABLED_FLAG", "y");
            ILD_DISPATCH_CORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
        }

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaEMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaJOB_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("JOB_CLASS", "Y");
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

        private void ILA_PAY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("PAY_GRADE", "Y");
        }

        private void ILA_POST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("POST", "Y");
        }

        private void ILA_DIR_INDIR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("DIR_INDIR_TYPE", "Y");
        }

        private void ILA_JOIN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            Set_Common_Parameter("JOIN", "Y");
        }


        #endregion

        #region ----- Cell Validating Event -----

        private void igrPERSON_INFO_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == IGR_PERSON.GetColumnToIndex("REPRE_NUM"))
            {
                if (Repre_Num_Validating_Check(e.NewValue) == false)
                {
                }
            }
        }

        #endregion

        #region ----- KeyDown Event -----

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