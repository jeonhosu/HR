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

namespace HRMF0108
{
    public partial class HRMF0108 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0108(Form pMainForm, ISAppInterface pAppInterface)
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
                return;
            }
            if (iString.ISNull(iedCLOSING_YYYYMM_0.EditValue) == string.Empty)
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10031"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            IDA_CLOSING.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_CLOSING.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_CLOSING.Fill();

            igrCLOSING.Focus();
        }

        #endregion

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
                    if (IDA_CLOSING.IsFocused)
                    {
                        IDA_CLOSING.AddOver();
                        igrCLOSING.SetCellValue("CORP_ID", iedCLOSING_TYPE_ID_0.EditValue);
                        igrCLOSING.SetCellValue("CORP_NAME", iedCLOSING_TYPE_NAME_0.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_CLOSING.IsFocused)
                    {
                        IDA_CLOSING.AddUnder();
                        igrCLOSING.SetCellValue("CORP_ID", iedCLOSING_TYPE_ID_0.EditValue);
                        igrCLOSING.SetCellValue("CORP_NAME", iedCLOSING_TYPE_NAME_0.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_CLOSING.IsFocused)
                    {
                        IDA_CLOSING.SetInsertParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                        IDA_CLOSING.SetInsertParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                        IDA_CLOSING.SetInsertParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                        IDA_CLOSING.SetUpdateParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);

                        IDA_CLOSING.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CLOSING.IsFocused)
                    {
                        IDA_CLOSING.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CLOSING.IsFocused)
                    {
                        IDA_CLOSING.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0108_Load(object sender, EventArgs e)
        {
            string Start_YYYYMM = "2010-01";
            string End_YYYYMM = iDate.ISYearMonth(DateTime.Today, 1);

            IDA_CLOSING.FillSchema();
            DefaultCorporation();

            ildCLOSING_YYYYMM.SetLookupParamValue("W_START_YYYYMM", Start_YYYYMM);
            ildCLOSING_YYYYMM.SetLookupParamValue("W_END_YYYYMM", End_YYYYMM);

            iedCLOSING_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            ildCLOSING_TYPE.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCLOSING_TYPE.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCLOSING_TYPE.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }

        private void ibtCREATION_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string sMessage = null;
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (string.IsNullOrEmpty(iedCLOSING_YYYYMM_0.EditValue.ToString()))
            {// 적용년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10031"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCLOSING_YYYYMM_0.Focus();
                return;
            }
            idcCLOSING_CREATE.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idcCLOSING_CREATE.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idcCLOSING_CREATE.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);

            //idcCLOSING_CREATE.SetCommandParamValue("O_MESSAGE", sMessage);
            idcCLOSING_CREATE.ExecuteNonQuery();
            sMessage = idcCLOSING_CREATE.GetCommandParamValue("O_MESSAGE").ToString();
            MessageBoxAdv.Show(sMessage, "Closing Create");
        }

        #endregion

        #region ----- Adapter Event -----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["CORP_ID"] == DBNull.Value)
            {// 업체
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CLOSING_YYYYMM"]) == string.Empty)
            {// 마감년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10031"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CLOSING_TYPE_ID"] == DBNull.Value)
            {// 마감항목
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10032"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ----- Lookup Event -----

        #endregion

    }
}