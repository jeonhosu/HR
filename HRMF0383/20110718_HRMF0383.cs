using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace HRMF0383
{
    public partial class HRMF0383 : Office2007Form
    {
        #region ----- Variables -----

        private ISCommonUtil.ISFunction.ISDateTime ISDate = new ISCommonUtil.ISFunction.ISDateTime();

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
        }

        #endregion;

        #region ----- MDi ToolBar Button Event -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchPerson();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaPERSON_INFO.IsFocused)
                    {
                        idaPERSON_INFO.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPERSON_INFO.IsFocused)
                    {
                        idaPERSON_INFO.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPERSON_INFO.IsFocused)
                    {
                        idaPERSON_INFO.Delete();
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

        #region ----- Private Method ----

        private void DefaultCorporation()
        {
            // 조회년월 SETTING
            ildYYYYMM_0.SetLookupParamValue("W_START_YYYYMM", "2010-01");

            WORK_YYYYMM_2.EditValue = ISDate.ISYearMonth(DateTime.Today);
            idcYYYYMM_TERM.SetCommandParamValue("W_YYYYMM", WORK_YYYYMM_2.EditValue);
            idcYYYYMM_TERM.ExecuteNonQuery();
            DATE_START_2.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_START_DATE");
            DATE_END_2.EditValue = idcYYYYMM_TERM.GetCommandParamValue("O_END_DATE");

            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            //작업장
            idcDEFAULT_FLOOR.ExecuteNonQuery();
            FLOOR_NAME_0.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_NAME");
            FLOOR_ID_0.EditValue = idcDEFAULT_FLOOR.GetCommandParamValue("O_FLOOR_ID");
            object oPERSON_NAME = idcDEFAULT_FLOOR.GetCommandParamValue("O_PERSON_NAME");
            object oCAPACITY = idcDEFAULT_FLOOR.GetCommandParamValue("O_CAPACITY"); //권한
            string vCAPACITY = ConvertString(oCAPACITY);
            
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
            EMPLOYE_TYPE_NAME_0.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE_NAME");
            EMPLOYE_TYPE_0.EditValue = idcDEFAULT_EMPLOYE_TYPE_0.GetCommandParamValue("O_CODE");
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void SearchPerson()
        {
            idaPERSON_INFO.Fill();
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0383_Load(object sender, EventArgs e)
        {
            DefaultCorporation();
            DefaultEmploye();

            idaPERSON_INFO.FillSchema();
        }

        #endregion;

        #region ----- LookUP Event ----

        private void ilaWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaWORK_TYPE_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildFLOOR_0.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildFLOOR_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaFLOOR_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildFLOOR_1.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildFLOOR_1.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("EMPLOYE_TYPE", "Y");
        }

        private void ilaYYYYMM_0_SelectedRowData(object pSender)
        {
            object vObject = WORK_YYYYMM_2.EditValue;
            string vYYYYMM = ConvertString(vObject);
            if (string.IsNullOrEmpty(vYYYYMM) == false)
            {
                System.DateTime v1stDate = ISDate.ISMonth_1st(vYYYYMM);
                System.DateTime vLastDate = ISDate.ISMonth_Last(vYYYYMM);
                DATE_START_2.EditValue = v1stDate.ToShortDateString();
                DATE_END_2.EditValue = vLastDate.ToShortDateString();
            }
        }

        #endregion
    }
}