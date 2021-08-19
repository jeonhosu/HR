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

namespace HRMF0104
{
    public partial class HRMF0104 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        public HRMF0104(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Data Find -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        } 

        private void SEARCH_DB()
        {
            if (iString.ISNull(Year_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Year_0.Focus();
                return;
            }
            if (iString.ISNull(TAX_TYPE_NAME_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Year_0.Focus();
                return;
            }
            IDA_TAX_RATE.Fill();
            igrTAX_RATE.Focus();
        }

        private void Insert_Data()
        {
            igrTAX_RATE.SetCellValue("TAX_YYYY", Year_0.EditValue);
            igrTAX_RATE.SetCellValue("TAX_TYPE_NAME", TAX_TYPE_NAME_0.EditValue);
            igrTAX_RATE.SetCellValue("TAX_TYPE_ID", TAX_TYPE_ID_0.EditValue);

            igrTAX_RATE.CurrentCellMoveTo(3);
        }

        #endregion

        #region ----- Application_MainButtonClick -----

        public void Application_MainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_TAX_RATE.IsFocused)
                    {
                        IDA_TAX_RATE.AddOver();
                        Insert_Data();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_TAX_RATE.IsFocused)
                    {
                        IDA_TAX_RATE.AddUnder();
                        Insert_Data();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_TAX_RATE.IsFocused)
                    {
                        IDA_TAX_RATE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_TAX_RATE.IsFocused)
                    {
                        IDA_TAX_RATE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_TAX_RATE.IsFocused)
                    {
                        IDA_TAX_RATE.Delete();
                    }
                }
            }
        }
        #endregion         

        #region ----- Form Event -----

        private void HRMF0104_Load(object sender, EventArgs e)
        {
            IDA_TAX_RATE.FillSchema();
        }

        private void HRMF0104_Shown(object sender, EventArgs e)
        {
            string Start_Year = iDate.ISYear(DateTime.Today, -10);
            string End_Year = iDate.ISYear(DateTime.Today, 1);

            ildYEAR.SetLookupParamValue("W_START_YEAR", Start_Year);
            ildYEAR.SetLookupParamValue("W_END_YEAR", End_Year);
        }

        private void ibtCOPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(Year_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Year_0.Focus();
                return;
            }
            if (iString.ISNull(TAX_TYPE_NAME_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                TAX_TYPE_NAME_0.Focus();
                return;
            }

            string mPre_YYYY = null;
            string mReturn_Value = null;
            string mSTATUS = "F";
            string mMESSAGE = null;
            DialogResult mDialogResult;

            // 전년도 자료 존재 체크
            mPre_YYYY = Convert.ToString(Convert.ToInt32(Year_0.EditValue) - Convert.ToInt32(1));
            idcTAX_RATE_CHECK_YN.SetCommandParamValue("W_TAX_YYYY", mPre_YYYY);
            idcTAX_RATE_CHECK_YN.ExecuteNonQuery();
            mReturn_Value = iString.ISNull(idcTAX_RATE_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));
            if (mReturn_Value == "N".ToString())
            {// 복사대상 없음 : 기존 자료 존재하지 않음.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10083"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Year_0.Focus();
                return;
            }

            // 당년도 자료 존재 체크
            idcTAX_RATE_CHECK_YN.SetCommandParamValue("W_TAX_YYYY", Year_0.EditValue);
            idcTAX_RATE_CHECK_YN.ExecuteNonQuery();
            mReturn_Value = iString.ISNull(idcTAX_RATE_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));
            if (mReturn_Value == "Y".ToString())
            {// 기존 자료 존재.
                mDialogResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (mDialogResult == DialogResult.No)
                {
                    return;
                }
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            // Copy 시작.
            idcTAX_RATE_COPY.ExecuteNonQuery();
            mSTATUS = iString.ISNull(idcTAX_RATE_COPY.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(idcTAX_RATE_COPY.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (idcTAX_RATE_COPY.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBoxAdv.Show(mMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            SEARCH_DB();
        }

        #endregion

        #region ----- Adapter Event -----

        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["TAX_YYYY"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:= Tax Rate Year(정산년도)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["TAX_TYPE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["START_AMOUNT"], 0) < Convert.ToInt32(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=End Amount(시작금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["END_AMOUNT"], 0) == Convert.ToInt32(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10025"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["END_AMOUNT"], 0) < iString.ISDecimaltoZero(e.Row["START_AMOUNT"], 0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10073"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISDecimaltoZero(e.Row["TAX_RATE"], 0) < Convert.ToInt32(0))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Tax Rate(세율)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {

        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildTAX_TYPE.SetLookupParamValue("W_GROUP_CODE", "TAX_TYPE");
            ildTAX_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaTAX_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildTAX_TYPE.SetLookupParamValue("W_GROUP_CODE", "TAX_TYPE");
            ildTAX_TYPE.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}