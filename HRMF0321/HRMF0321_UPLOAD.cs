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

namespace HRMF0321
{
    public partial class HRMF0321_UPLOAD : Office2007Form
    {
        
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string gSTATUS = "N";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.

        #endregion;

        #region ----- Constructor -----

        public HRMF0321_UPLOAD(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            gSTATUS = "N";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.
        }

        #endregion;

        #region ----- Private Methods ----



        #endregion;

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

        #region ----- Excel Upload : Asset Master -----

        private void Select_Excel_File()
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FILE_PATH.EditValue = openFileDialog1.FileName;
                }
                else
                {
                    FILE_PATH.EditValue = string.Empty;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        private void Excel_Upload()
        {
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;
            bool vXL_Load_OK = false;

            if (iString.ISNull(FILE_PATH.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FILE_PATH))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vOPenFileName = FILE_PATH.EditValue.ToString();
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
                return;
            }


            //기존자료 삭제.
            vSTATUS = "F";
            vMESSAGE = string.Empty;

            IDC_DELETE_EXCEL_UPLOADING.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_DELETE_EXCEL_UPLOADING.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_DELETE_EXCEL_UPLOADING.GetCommandParamValue("O_MESSAGE"));
            if (IDC_DELETE_EXCEL_UPLOADING.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 업로드 아답터 fill //
            IDA_EXCEL_UPLOADING.Fill();

            try
            {
                if (vXL_Load_OK == true)
                {
                    vXL_Load_OK = vXL_Upload.LoadXL(IDA_EXCEL_UPLOADING, 2);
                    if (vXL_Load_OK == false)
                    {
                        IDA_EXCEL_UPLOADING.Cancel();
                    }
                    else
                    {
                        IDA_EXCEL_UPLOADING.Update();
                    }
                }
            }
            catch (Exception ex)
            {
                IDA_EXCEL_UPLOADING.Cancel();
                isAppInterfaceAdv1.OnAppMessage(ex.Message);

                vXL_Upload.DisposeXL();

                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();
                return;
            }
            vXL_Upload.DisposeXL();


            if (IDA_EXCEL_UPLOADING.IsUpdateCompleted == true)
            {
                IDA_EXCEL_UPLOADED.Fill();
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Set_Trans_Work_Calendar()
        {
            string vSTATUS = string.Empty;
            string vMESSAGE = string.Empty;

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            vSTATUS = "F";
            vMESSAGE = string.Empty;

            IDC_TRANS_WORK_CALENDAR.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_TRANS_WORK_CALENDAR.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_TRANS_WORK_CALENDAR.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (IDC_TRANS_WORK_CALENDAR.ExcuteError || vSTATUS == "F")
            {
                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            IDA_EXCEL_UPLOADED.Fill();
        }

        #endregion

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0321_UPLOAD_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0321_UPLOAD_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (gSTATUS == "P")
            {
                e.Cancel = true;
                return;
            }
        }

        private void BTN_FILE_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File();
        }

        private void BTN_UPLOAD_EXEC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            gSTATUS = "P";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.
            Excel_Upload();
            gSTATUS = "N";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.
        }

        private void BTN_TRANS_WORK_CALENDAR_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            gSTATUS = "P";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.
            Set_Trans_Work_Calendar();
            gSTATUS = "N";  //P-Processing, N-Normal 체크해서 Processing중일 경우 닫기 제어.
        }

        private void BTN_CLOSED_FORM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (gSTATUS == "P")
            {
                return;
            }
            this.Close();
        }

        #endregion              

        #region ----- Lookup Event -----

        #endregion

    }
}