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

namespace HRMF0602
{
    public partial class HRMF0602 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----
        public HRMF0602(Form pMainForm, ISAppInterface pAppInterface)
        {
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
            if (CORP_ID_0.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            if (iString.ISNull(ADJUSTMENT_TYPE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ADJUSTMENT_TYPE_NAME_0.Focus();
                return;
            }

            IDA_RETIRE_INSURANCE.FillSchema();
            igrRETIRE_INSURANCE.Focus();
        }

        private void Init_Insert()
        {
            igrRETIRE_INSURANCE.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            igrRETIRE_INSURANCE.SetCellValue("ADJUSTMENT_TYPE_NAME", ADJUSTMENT_TYPE_NAME_0.EditValue);
            igrRETIRE_INSURANCE.SetCellValue("ADJUSTMENT_TYPE", ADJUSTMENT_TYPE_0.EditValue);
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
                    if (IDA_RETIRE_INSURANCE.IsFocused)
                    {
                        IDA_RETIRE_INSURANCE.AddOver();
                    }
                    Init_Insert();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_RETIRE_INSURANCE.IsFocused)
                    {
                        IDA_RETIRE_INSURANCE.AddUnder();
                    }
                    Init_Insert();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_RETIRE_INSURANCE.Update();                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_RETIRE_INSURANCE.IsFocused)
                    {
                        IDA_RETIRE_INSURANCE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_RETIRE_INSURANCE.IsFocused)
                    {
                        IDA_RETIRE_INSURANCE.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0602_Load(object sender, EventArgs e)
        {
            IDA_RETIRE_INSURANCE.FillSchema();
            DefaultCorporation();              //Default Corp.
        }
        #endregion  

        #region ----- Adapter Event -----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["ADJUSTMENT_TYPE"]) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ADJUSTMENT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10071"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["INSUR_CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10072"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {

        }
        #endregion

        #region ----- LookUp Event -----
        private void ilaADJUSTMENT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "RETIRE_ADJUSTMENT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaINSUR_CORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "INSUR_CORP");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }
        
        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON_0.SetLookupParamValue("W_START_DATE", iDate.ISMonth_1st(DateTime.Today, -DateTime.Today.Month));
            ildPERSON_0.SetLookupParamValue("W_END_DATE", DateTime.Today);
        }

        private void ilaINSUR_CORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "INSUR_CORP");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }
        
        #endregion

    }
}