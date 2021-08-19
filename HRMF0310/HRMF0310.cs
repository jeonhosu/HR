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

namespace HRMF0310
{
    public partial class HRMF0310 : Office2007Form
    {
        ISFunction.ISDateTime iSDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0310(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

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
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_FR_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_FR_0.Focus();
                return;
            }
                        
            idaDAY_INTERFACE.Fill();
            igrDAY_INTERFACE.Focus();
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = WORK_DATE_FR_0.EditValue;

            idaWORK_CALENDAR.SetSelectParamValue("W_END_DATE", pWork_Date);
            idaDAY_HISTORY.Fill();
            idaDUTY_PERIOD.Fill();
            idaWORK_CALENDAR.Fill();
        }

        private void isSearch_Day_History(int pAdd_Day)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(WORK_DATE_8.EditValue) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = Convert.ToDateTime(WORK_DATE_8.EditValue).AddDays(pAdd_Day);
            idaDAY_HISTORY.Fill();
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
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaDAY_INTERFACE.IsFocused)
                    {
                        idaDAY_INTERFACE.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0310_Load(object sender, EventArgs e)
        {
            idaDAY_INTERFACE.FillSchema();
            WORK_DATE_FR_0.EditValue = DateTime.Today;
            WORK_DATE_TO_0.EditValue = DateTime.Today;

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
            irbAPPR_A.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EMAIL_STATUS.EditValue = "N";
        }

        private void irbSTATUS_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv isINOUT = sender as ISRadioButtonAdv;
            APPROVE_STATUS_0.EditValue = isINOUT.RadioCheckedString;

            // refill.
            Search_DB();
        }

        private void ibtOK_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 승인
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_OK";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_OK";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            idaDAY_INTERFACE.SetUpdateParamValue("P_APPROVE_FLAG", "OK");
            idaDAY_INTERFACE.SetUpdateParamValue("P_APPROVE_STATUS", APPROVE_STATUS_0.EditValue);
            idaDAY_INTERFACE.Update();

            // refill.
            Search_DB();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 취소
            // EMAIL STATUS.
            if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "A".ToString())
            {
                EMAIL_STATUS.EditValue = "A_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "B".ToString())
            {
                EMAIL_STATUS.EditValue = "B_CANCEL";
            }
            else if (iString.ISNull(APPROVE_STATUS_0.EditValue) == "C".ToString())
            {
                EMAIL_STATUS.EditValue = "C_CANCEL";
            }
            else
            {
                EMAIL_STATUS.EditValue = "N";
            }
            idaDAY_INTERFACE.SetUpdateParamValue("P_APPROVE_FLAG", "CANCEL");
            idaDAY_INTERFACE.SetUpdateParamValue("P_APPROVE_STATUS", APPROVE_STATUS_0.EditValue);
            idaDAY_INTERFACE.Update();

            // refill.
            Search_DB();
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }

        private void igrDAY_INTERFACE_CellDoubleClick(object pSender)
        {
            if (igrDAY_INTERFACE.RowIndex < 0 && igrDAY_INTERFACE.ColIndex == igrDAY_INTERFACE.GetColumnToIndex("CHECK_YN"))
            {
                for (int r = 0; r < igrDAY_INTERFACE.RowCount; r++)
                {
                    if (iString.ISNull(igrDAY_INTERFACE.GetCellValue(r, igrDAY_INTERFACE.GetColumnToIndex("CHECK_YN")), "N") == "Y".ToString())
                    {
                        igrDAY_INTERFACE.SetCellValue(r, igrDAY_INTERFACE.GetColumnToIndex("CHECK_YN"), "N");
                    }
                    else
                    {
                        igrDAY_INTERFACE.SetCellValue(r, igrDAY_INTERFACE.GetColumnToIndex("CHECK_YN"), "Y");
                    }
                }
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation Name(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDAY_INTERFACE_PreDelete(ISPreDeleteEventArgs e)
        {            
        }

        private void idaDAY_INTERFACE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            isSearch_WorkCalendar(igrDAY_INTERFACE.GetCellValue("PERSON_ID"), igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
        }

        private void idaDAY_INTERFACE_UpdateCompleted(object pSender)
        {
            // EMAIL 발송.
            idcEMAIL_SEND.SetCommandParamValue("P_GUBUN", EMAIL_STATUS.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_SOURCE_TYPE", "WORK");
            idcEMAIL_SEND.SetCommandParamValue("P_CORP_ID", CORP_ID_0.EditValue);
            idcEMAIL_SEND.SetCommandParamValue("P_WORK_DATE", DateTime.Today);
            idcEMAIL_SEND.SetCommandParamValue("P_REQ_DATE", DateTime.Today);
            idcEMAIL_SEND.ExecuteNonQuery();
        }

        #endregion

        #region ----- LookUp Event -----
        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }
        
        private void ildWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", WORK_DATE_FR_0.EditValue);
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_MODIFY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }        
        #endregion

        

    }
}