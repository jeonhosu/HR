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

namespace HRMF0320
{
    public partial class HRMF0320 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0320()
        {
            InitializeComponent();
        }

        public HRMF0320(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (string.IsNullOrEmpty(STD_DATE_0.EditValue.ToString()))
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE_0.Focus();
                return;
            }

            idaDUTY_EXCEPTION.Fill();
            gridDUTY_EXCEPTION.Focus();
        }

        private void AddLine() //레코드 추가 시, Setting
        {
            gridDUTY_EXCEPTION.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            gridDUTY_EXCEPTION.SetCellValue("ENABLED_FLAG", "Y");        
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        #endregion;

        #region ----- Events -----

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
                    idaDUTY_EXCEPTION.AddOver();
                    AddLine();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    idaDUTY_EXCEPTION.AddUnder();
                    AddLine();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaDUTY_EXCEPTION.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    idaDUTY_EXCEPTION.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    idaDUTY_EXCEPTION.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (idaDUTY_EXCEPTION.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (idaDUTY_EXCEPTION.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;        

        #region ----- Form event -----

        private void HRMF0320_Load(object sender, EventArgs e)
        {
            // FillSchema
            idaDUTY_EXCEPTION.FillSchema();            
        }

        private void HRMF0320_Shown(object sender, EventArgs e)
        {
            STD_DATE_0.EditValue = DateTime.Today;
            DefaultCorporation();
            ENABLED_YN.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        #endregion

        #region ----- Lookup Event -----

        // 작업장 Lookup
        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        // 성명 Lookup
        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        // 직위 Lookup
        private void ilaPOST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter event -----

        private void idaDUTY_EXCEPTION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=[sData]"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaDUTY_EXCEPTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=[Person No]"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //조정 시간//
            if (iString.ISNull(e.Row["ADJUST_WORKTIME_YN"]) == "Y")
            {
                if (iString.ISNumtoZero(e.Row["IN_TIME"]) == 0 && iString.ISNumtoZero(e.Row["OUT_TIME"]) == 0)
                {
                    MessageBoxAdv.Show("Adjust time checked => Insert time must have value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            else
            {
                if (iString.ISNumtoZero(e.Row["IN_TIME"]) != 0 || iString.ISNumtoZero(e.Row["OUT_TIME"]) != 0)
                {
                    MessageBoxAdv.Show("Adjust time not checked => Insert time is not value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }

            //출근시간 = 고정//
            if (iString.ISNull(e.Row["AUTO_WORKTIME_IN_YN"]) == "Y")
            {
                if (iString.ISNull(e.Row["DAY_FIX_IN_TIME"]) == string.Empty && iString.ISNull(e.Row["NIGHT_FIX_IN_TIME"]) == string.Empty)
                {
                    MessageBoxAdv.Show("Fix in time checked => Insert time must have value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            } 
            if (iString.ISNull(e.Row["AUTO_WORKTIME_OUT_YN"]) == "Y")
            {
                if (iString.ISNull(e.Row["DAY_FIX_OUT_TIME"]) == string.Empty && iString.ISNull(e.Row["NIGHT_FIX_OUT_TIME"]) == string.Empty)
                {
                    MessageBoxAdv.Show("Fix Out time checked => Insert time must have value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            } 
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_TO"]) != string.Empty)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자~종료일자
                    e.Cancel = true;
                    return;
                }
            }
             
            //검증//
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_PERSON_ID", e.Row["PERSON_ID"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_ADJUST_WORKTIME_YN", e.Row["ADJUST_WORKTIME_YN"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_IN_TIME", e.Row["IN_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_OUT_TIME", e.Row["OUT_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_AUTO_WORKTIME_IN_YN", e.Row["AUTO_WORKTIME_IN_YN"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_AUTO_WORKTIME_OUT_YN", e.Row["AUTO_WORKTIME_OUT_YN"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_DAY_FIX_IN_ADD_DAY", e.Row["DAY_FIX_IN_ADD_DAY"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_DAY_FIX_IN_TIME", e.Row["DAY_FIX_IN_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_DAY_FIX_OUT_ADD_DAY", e.Row["DAY_FIX_OUT_ADD_DAY"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_DAY_FIX_OUT_TIME", e.Row["DAY_FIX_OUT_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_NIGHT_FIX_IN_ADD_DAY", e.Row["NIGHT_FIX_IN_ADD_DAY"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_NIGHT_FIX_IN_TIME", e.Row["NIGHT_FIX_IN_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_NIGHT_FIX_OUT_ADD_DAY", e.Row["NIGHT_FIX_OUT_ADD_DAY"]);
            IDC_VALIDATE_DUTY_EXCEPTION.SetCommandParamValue("P_NIGHT_FIX_OUT_TIME", e.Row["NIGHT_FIX_OUT_TIME"]);
            IDC_VALIDATE_DUTY_EXCEPTION.ExecuteNonQuery();
            string vSTATUS = iString.ISNull(IDC_VALIDATE_DUTY_EXCEPTION.GetCommandParamValue("O_STATUS"));
            string vMESSAGE = iString.ISNull(IDC_VALIDATE_DUTY_EXCEPTION.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                e.Cancel = true;
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);  // 시작일자~종료일자
                }
                return;
            }
        }

        #endregion
    }
}