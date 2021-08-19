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

namespace HRMF0109
{
    public partial class HRMF0109 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        public HRMF0109(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property / Method ----
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
            if (iString.ISNull(iedSTD_DATE_0.EditValue) == string.Empty)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_PERSON_CC.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_PERSON_CC.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            IDA_PERSON_CC.Fill();
            igrHRM_PERSON_CC_G.Focus();
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----
        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_PERSON_CC.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_PERSON_CC.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_PERSON_CC.IsFocused)
                    {
                        IDA_PERSON_CC.SetUpdateParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);

                        IDA_PERSON_CC.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_PERSON_CC.IsFocused)
                    {
                        IDA_PERSON_CC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_PERSON_CC.IsFocused)
                    {
                        IDA_PERSON_CC.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0109_Load(object sender, EventArgs e)
        {
            // FillSchema
            IDA_PERSON_CC.FillSchema();

            iedSTD_DATE_0.EditValue = DateTime.Today;
            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }

        #endregion

        #region ----- Data Adapter Event ----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["FLOOR_ID"] == DBNull.Value)
            //{// FLOOR_ID
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10017"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (e.Row["CC_ID"] == DBNull.Value)
            //{// cc
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10018"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}

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
        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCOMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCOMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ILA_DIR_INDIR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //DIR/INDIR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DIR_INDIR_TYPE");
            ildCOMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCOMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaEMPLOYE_0_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //EMPLOYE
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCOMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // DEPT
            ildDEPT.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildDEPT.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // PERSON
            ildPERSON.SetLookupParamValue("W_PERSON_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildPERSON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }
        private void ilaCOST_CENTER_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaCOST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        #endregion

        
    }
}