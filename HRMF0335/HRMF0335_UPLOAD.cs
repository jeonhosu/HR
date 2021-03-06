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
using System.IO;

namespace HRMF0335
{
    public partial class HRMF0335_UPLOAD : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        int mERROR_COUNT = 0;

        #region ----- Constructor -----

        public HRMF0335_UPLOAD(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            //this.MdiParent = pMainForm;
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

        #endregion

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Excel Upload -----

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    UPLOAD_FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    UPLOAD_FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private bool Excel_Upload()
        {
            bool vResult = false;
            
            if (iString.ISNull(UPLOAD_FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(UPLOAD_FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vResult;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            bool vXL_Load_OK = false;
            string vOPenFileName = UPLOAD_FILE_PATH.EditValue.ToString();
            XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);
            try
            {
                vXL_Upload.OpenFileName = vOPenFileName;
                vXL_Load_OK = vXL_Upload.OpenXL();
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return vResult;
            }

            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDA_WORK_SCHEDULE_UPLOAD, 2);
                    if (vXL_Load_OK == false)
                    {
                        IDA_WORK_SCHEDULE_UPLOAD.Cancel();
                        vResult = false;
                    }
                    else
                    {
                        IDA_WORK_SCHEDULE_UPLOAD.Update();
                        vResult = true;
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_WORK_SCHEDULE_UPLOAD.Cancel();
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                vResult = false;
                return vResult;
            }
            vXL_Upload.DisposeXL();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            
            return vResult;
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_WORK_SCHEDULE_UPLOAD.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_WORK_SCHEDULE_UPLOAD.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_WORK_SCHEDULE_UPLOAD.IsFocused)
                    {
                        IDA_WORK_SCHEDULE_UPLOAD.SetUpdateParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);

                        IDA_WORK_SCHEDULE_UPLOAD.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_WORK_SCHEDULE_UPLOAD.IsFocused)
                    {
                        IDA_WORK_SCHEDULE_UPLOAD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0335_UPLOAD_Load(object sender, EventArgs e)
        {
            // FillSchema
            IDA_WORK_SCHEDULE_UPLOAD.FillSchema();
        }

        private void HRMF0335_UPLOAD_Shown(object sender, EventArgs e)
        {
            
        }

        private void BTN_SELECT_EXCEL_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_FILE_UPLOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_WORK_SCHEDULE_UPLOAD.Fill();
            if (Excel_Upload() == true)
            {
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                
                this.DialogResult = DialogResult.Cancel;
            }
            this.Close();
        }

        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_WORK_SCHEDULE_UPLOAD.Cancel();
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void IDA_WORK_SCHEDULE_UPLOAD_UpdateCompleted(object pSender)
        {
            int vIDX_STATUS = IGR_WORK_SCHEDULE.GetColumnToIndex("STATUS");
            int vIDX_MESSAGE = IGR_WORK_SCHEDULE.GetColumnToIndex("MESSAGE");
            string vMESSAGE = String.Empty;
            // 저장 완료후 오류가 있으면 업로드 화면을 닫지 않고 놔둡니다.
            for (int vROW = 0; vROW < IGR_WORK_SCHEDULE.RowCount; vROW++)
            {
                if ("F" == iString.ISNull(IGR_WORK_SCHEDULE.GetCellValue(vROW, vIDX_STATUS)))
                {
                    mERROR_COUNT++;
                    if (vMESSAGE == String.Empty)
                    {
                        vMESSAGE = iString.ISNull(IGR_WORK_SCHEDULE.GetCellValue(vROW, vIDX_MESSAGE));
                    }
                }
            }
            if (mERROR_COUNT > 0)
            {
                // 오류가 있을 경우 화면을 닫지 않음 //
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(string.Format("{0}-{1}",isMessageAdapter1.ReturnText("FCM_10238", string.Format("&&VALUE:=[{0}]", mERROR_COUNT)), vMESSAGE), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion


    }
}