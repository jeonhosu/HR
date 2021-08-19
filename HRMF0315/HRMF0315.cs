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

namespace HRMF0315
{
    public partial class HRMF0315 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----
        public HRMF0315(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            string iYear;

            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }            
            if (STD_YYYYMM_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYYMM_0.Focus();
                return;
            }
            iYear = Convert.ToDateTime(END_DATE_0.EditValue).Year.ToString();
            idaHOLIDAY_MANAGEMENT.SetSelectParamValue("W_STD_YEAR", iYear);
            idaHOLIDAY_DUTY_SUMMARY.SetSelectParamValue("W_STD_YEAR", iYear);
            idaHOLIDAY_MANAGEMENT.Fill();
            igrHOLIDAY_MANAGEMENT.Focus();
        }

        private void isInit_Num(string pHoliday_Type)
        {
            int Pre_Next_num = 0;
            int Creation_num = 0;
            int Plus_num = 0;
            int Use_num = 0;
            int Total_Creation_num = 0;

            if (pHoliday_Type == "1".ToString())
            {
                if (NY_PRE_NEXT_NUM.EditValue != null)
                {
                    Pre_Next_num = Convert.ToInt32(NY_PRE_NEXT_NUM.EditValue);
                }
                if (NY_CREATION_NUM.EditValue != null)
                {
                    Creation_num = Convert.ToInt32(NY_CREATION_NUM.EditValue);
                }
                if (NY_PLUS_NUM.EditValue != null)
                {
                    Plus_num = Convert.ToInt32(NY_PLUS_NUM.EditValue);
                }
                if (NY_PLUS_NUM.EditValue != null)
                {
                    Use_num = Convert.ToInt32(NY_PLUS_NUM.EditValue);
                }
                Total_Creation_num = Pre_Next_num + Creation_num + Plus_num;
                NY_TOTAL_CREATION_NUM.EditValue = Total_Creation_num;
                NY_REMAIN_NUM.EditValue = Total_Creation_num - Use_num;
            }
            else if (pHoliday_Type == "2".ToString())
            {
                if (SM_CREATION_NUM.EditValue != null)
                {
                    Creation_num = Convert.ToInt32(SM_CREATION_NUM.EditValue);
                }
                if (SM_USE_NUM.EditValue != null)
                {
                    Use_num = Convert.ToInt32(SM_USE_NUM.EditValue);
                }
                SM_REMAIN_NUM.EditValue = Creation_num - Use_num;
            }
            else if (pHoliday_Type == "3".ToString())
            {
                if (SP_CREATION_NUM.EditValue != null)
                {
                    Creation_num = Convert.ToInt32(SP_CREATION_NUM.EditValue);
                }
                if (SP_USE_NUM.EditValue != null)
                {
                    Use_num = Convert.ToInt32(SP_USE_NUM.EditValue);
                }
                SP_REMAIN_NUM.EditValue = Creation_num - Use_num;
            }
        }

        private void Init_Period_Date()
        {
            if (iString.ISNull(STD_YYYYMM_0.EditValue) == string.Empty)
            {
                return;
            }
            DateTime mSTD_DATE = DateTime.Parse(string.Format("{0}{1}", STD_YYYYMM_0.EditValue, "-01"));
            START_DATE_0.EditValue = iDate.ISMonth_1st(mSTD_DATE);
            END_DATE_0.EditValue = iDate.ISMonth_Last(mSTD_DATE);
        }

        private void Show_Execute_Duty(object pExecute_Type)
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iString.ISNull(STD_YYYYMM_0.EditValue) == string.Empty)
            {// 월근태년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYYMM_0.Focus();
                return;
            }

            object mExecute_Type = pExecute_Type;
            DialogResult vdlgResult;            
            Form vEXECUTE_DUTY = new HRMF0315_DUTY(isAppInterfaceAdv1.AppInterface, mExecute_Type
                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                , DEPT_ID_0.EditValue, DEPT_NAME_0.EditValue
                , PERSON_ID_0.EditValue, NAME_0.EditValue
                , END_DATE_0.EditValue);
            vdlgResult = vEXECUTE_DUTY.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
            }
            vEXECUTE_DUTY.Dispose();
        }
        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        
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
                    if (idaHOLIDAY_MANAGEMENT.IsFocused)
                    {
                        string iYear;

                        iYear = Convert.ToDateTime(END_DATE_0.EditValue).Year.ToString();
                        idaHOLIDAY_MANAGEMENT.SetUpdateParamValue("W_DUTY_YEAR", iYear);
                        idaHOLIDAY_MANAGEMENT.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHOLIDAY_MANAGEMENT.IsFocused)
                    {
                        idaHOLIDAY_MANAGEMENT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHOLIDAY_MANAGEMENT.IsFocused)
                    {
                        idaHOLIDAY_MANAGEMENT.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0315_Load(object sender, EventArgs e)
        {   
            this.Visible = true;
            idaHOLIDAY_MANAGEMENT.FillSchema();

            // Year Month Setting
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2000-01");
            STD_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            Init_Period_Date();
                        
            // CORP SETTING
            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
        }

        private void ibtSET_HOLIDAY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Execute_Duty("HOLIDAY");
        }

        private void ibtSET_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Execute_Duty("PAYMENT");
        }

        private void isTRANS_PAYMENT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Show_Execute_Duty("TRANSFER");
        }

        private void NY_NUM_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            isInit_Num("1");
        }

        private void SM_CREATION_NUM_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            isInit_Num("2");
        }

        private void SP_CREATION_NUM_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            isInit_Num("3");
        }
        #endregion  

        #region ----- Adapter Event -----
        private void idaMONTH_TOTAL_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Row["STD_YEAR"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Info(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            
        }
        #endregion

        #region ----- LookUp Event -----
        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", END_DATE_0.EditValue);
        }
        #endregion

    }
}