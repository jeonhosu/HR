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

namespace HRMF0631
{
    public partial class HRMF0631 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Constructor -----
        public HRMF0631(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void SEARCH_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (W_STD_DATE.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IGR_BANK_TRANSFER.LastConfirmChanges();
            IDA_BANK_TRANSFER.OraSelectData.AcceptChanges();
            IDA_BANK_TRANSFER.Refillable = true;

            IDA_BANK_TRANSFER.Fill();
            IGR_BANK_TRANSFER.Focus();
        }
         
        private void InsertDB()
        {
            IGR_BANK_TRANSFER.Focus();
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
                    if (IDA_BANK_TRANSFER.IsFocused)
                    {
                        IDA_BANK_TRANSFER.AddOver();
                        InsertDB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_BANK_TRANSFER.IsFocused)
                    {
                        IDA_BANK_TRANSFER.AddUnder();
                        InsertDB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_BANK_TRANSFER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_BANK_TRANSFER.Cancel(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if(IDA_BANK_TRANSFER.CurrentRow.RowState == DataRowState.Added)
                    {
                        IDA_BANK_TRANSFER.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0631_Load(object sender, EventArgs e)
        { 
        }

        private void HRMF0631_Shown(object sender, EventArgs e)
        {
            W_STD_DATE.EditValue = iDate.ISYearMonth(DateTime.Today);
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]

            DefaultCorporation();                  // Corp Default Value Setting. 
            // FillSchema
            IDA_BANK_TRANSFER.FillSchema();
        }
           

        private void BTN_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0631_UPLOAD vHRMF0631_UPLOAD = new HRMF0631_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface, W_CORP_ID.EditValue, W_STD_DATE.EditValue);
            vdlgResult = vHRMF0631_UPLOAD.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            vHRMF0631_UPLOAD.Dispose();
            SEARCH_DB();
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void IDA_BANK_TRANSFER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void IDA_BANK_TRANSFER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        } 

        #endregion

        #region ----- Lookup Event -----

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildDEPT_0.SetLookupParamValue("W_USABLE_CHECK_YN", "Y"); 
        }

        private void ILA_POST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        } 

        private void ilaPERSON_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", W_DEPT_ID.EditValue);
            ildPERSON.SetLookupParamValue("W_POST_ID", W_POST_ID.EditValue);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", W_FLOOR_ID.EditValue);
        }

        private void ILA_PERSON_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_DEPT_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_POST_ID", DBNull.Value);
            ildPERSON.SetLookupParamValue("W_FLOOR_ID", DBNull.Value);
        }
        #endregion

    }
}