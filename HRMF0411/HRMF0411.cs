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

namespace HRMF0411
{
    public partial class HRMF0411 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Constructor -----
        public HRMF0411(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
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
            if (INSUR_YYYYMM_0.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_INSUR_MASTER.Fill();
            igrINSURANCE.Focus();
        }

        private void Set_CheckBox()
        {
            int mIDX_Col = igrINSURANCE.GetColumnToIndex("SOCIAL_INSUR_YN");
            int mIDX_EFFECTIVE_FR = igrINSURANCE.GetColumnToIndex("INSUR_YYYYMM_FR");

            object mCheck_YN = CB_SELECT.CheckBoxValue;
            for (int r = 0; r < igrINSURANCE.RowCount; r++)
            {
                igrINSURANCE.SetCellValue(r, mIDX_Col, mCheck_YN);
                if (mCheck_YN.ToString() == "Y")
                {
                    if (iString.ISNull(igrINSURANCE.GetCellValue(r, mIDX_EFFECTIVE_FR)) == string.Empty)
                    {
                        igrINSURANCE.SetCellValue(r, mIDX_EFFECTIVE_FR, INSUR_YYYYMM_0.EditValue);
                    }
                }
                //else
                //{
                //    igrINSURANCE.SetCellValue(r, mIDX_EFFECTIVE_FR, string.Empty);
                //}
            }
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
                    if (IDA_INSUR_MASTER.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_INSUR_MASTER.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_INSUR_MASTER.IsFocused)
                    {
                        IDA_INSUR_MASTER.SetUpdateParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);

                        IDA_INSUR_MASTER.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_INSUR_MASTER.IsFocused)
                    {
                        IDA_INSUR_MASTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0411_Load(object sender, EventArgs e)
        {
            // FillSchema
            IDA_INSUR_MASTER.FillSchema();
        }

        private void HRMF0411_Shown(object sender, EventArgs e)
        {
            INSUR_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]

            DefaultCorporation();                  // Corp Default Value Setting.
            RB_UNENROLLED.CheckedState = ISUtil.Enum.CheckedState.Checked;
            INSUR_STATUS_9.EditValue = "N";
        }

        private void RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            INSUR_STATUS_9.EditValue = iStatus.RadioButtonString;

            SEARCH_DB();
        }

        private void CB_SELECT_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            Set_CheckBox();
        }

        private void igrINSURANCE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.ColIndex == igrINSURANCE.GetColumnToIndex("SOCIAL_INSUR_YN"))
            {
                if (iString.ISNull(e.NewValue) == "Y" && iString.ISNull(igrINSURANCE.GetCellValue("INSUR_YYYYMM_FR")) == string.Empty)
                {
                    igrINSURANCE.SetCellValue("INSUR_YYYYMM_FR", INSUR_YYYYMM_0.EditValue);
                }
                //else if (e.NewValue != "Y" && iString.ISNull(igrINSURANCE.GetCellValue("INSUR_YYYYMM_FR")) != string.Empty)
                //{
                //    igrINSURANCE.SetCellValue("INSUR_YYYYMM_FR", string.Empty);
                //}
            }
        }

        private void BTN_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0411_UPLOAD vHRMF0411_UPLOAD = new HRMF0411_UPLOAD(this.MdiParent, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0411_UPLOAD.ShowDialog();
            if (vdlgResult == DialogResult.Cancel)
            {
                return;
            }
            vHRMF0411_UPLOAD.Dispose();
            SEARCH_DB();
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SOCIAL_INSUR_YN"]) == "Y" && iString.ISNull(e.Row["INSUR_YYYYMM_FR"]) == string.Empty)
            {// cc
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["SOCIAL_INSUR_YN"]) == "N" && iString.ISNull(e.Row["INSUR_YYYYMM_FR"]) != string.Empty)
            {// cc
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10464"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
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

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_EMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        
    }
}