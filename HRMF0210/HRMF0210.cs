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

namespace HRMF0210
{
    public partial class HRMF0210 : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();



        #endregion;

        #region ----- Constructor -----

        public HRMF0210()
        {
            InitializeComponent();
        }

        public HRMF0210(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (iString.ISNull(CORP_ID_0.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        CORP_NAME_0.Focus();
                        return;
                    }
                    if (iString.ISNull(STANDARD_DATE_0.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        STANDARD_DATE_0.Focus();
                        return;
                    }
                    idaLENGTH_ONE_SRV.Fill();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                        idaLENGTH_ONE_SRV.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                        idaLENGTH_ONE_SRV.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                        idaLENGTH_ONE_SRV.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                        idaLENGTH_ONE_SRV.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                        idaLENGTH_ONE_SRV.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if (idaLENGTH_ONE_SRV.IsFocused)
                    {
                    }
                }
            }
        }

        #endregion;

       

        private void HRMF0210_Load_1(object sender, EventArgs e)
        {
            // Lookup SETTING
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            // Standard Date SETTING
            STANDARD_DATE_0.EditValue = DateTime.Today;

        }

    }
}