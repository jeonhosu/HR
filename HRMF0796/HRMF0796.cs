using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;       //호환되지 않은DLL을 사용할 때.

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0796
{
    public partial class HRMF0796 : Office2007Form
    {
        #region ----- API Dll Import -----

        [DllImport("fcrypt_es.dll")]
        extern public static int DSFC_EncryptFile(int hWnd, string pszPlainFilePathName, string pszEncFilePathName, string pszPassword, uint nOption);

        string inputPath;
        string OutputPath;
        string Password;
        uint DSFC_OPT_OVERWRITE_OUTPUT;
        int nRet;

        #endregion;
        
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0796()
        {
            InitializeComponent();
        }

        public HRMF0796(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;            
        }

        #endregion;

        #region ----- Private Methods ----
        
        private void DefaultCorporation()
        {
            try
            {
                // Lookup SETTING
                ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
                ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

                // LOOKUP DEFAULT VALUE SETTING - CORP
                IDC_DEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
                IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                IDC_DEFAULT_CORP.ExecuteNonQuery();
                W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

                W_CORP_NAME.BringToFront();
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void SearchDB()
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_YEAR_YYYY.EditValue) == string.Empty)
            {
                W_YEAR_YYYY.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_YEAR_YYYY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(W_HALF_TYPE.EditValue) == string.Empty)
            {
                W_HALF_TYPE_NAME.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_HALF_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(W_PERIOD_FR.EditValue) == string.Empty)
            {
                W_PERIOD_FR.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(W_PERIOD_TO.EditValue) == string.Empty)
            {
                W_PERIOD_TO.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_PERIOD_TO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(W_HALF_SEQ.EditValue) == string.Empty)
            {
                W_HALF_SEQ_TYPE.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_HALF_SEQ_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }  

            if (TB_MAIN.SelectedTab.TabIndex == TP_E_FILE.TabIndex)
            {
                IDA_eFILE_INFO.Fill();
                IDA_FILE_SUM.Fill();
            }
            else if(TB_MAIN.SelectedTab.TabIndex == TP_FILE_DTL.TabIndex)
            {
                SearchDB_Sub();
            }
            else
            {  
                IDA_BSN_INCOME.Fill();
                IGR_PAYMENT_STATEMENT.Focus();
            }
        }

        private void SearchDB_Sub()
        {
            IDA_FILE_DTL.Fill();
        }

        private void SetParameter(object pGROUP_CODE, object pENABLED_FLAG)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGROUP_CODE);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pENABLED_FLAG);
        }

        private string EXPORT_VALIDATE()
        {
            string vRETURN = "N";
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }

            if (iConv.ISNull(START_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(END_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (Convert.ToDateTime(START_DATE.EditValue) > Convert.ToDateTime(END_DATE.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(NAME.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(NAME)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(TEL_NUMBER.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(TEL_NUMBER)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(TAX_PROGRAM_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(TAX_PROGRAM_CODE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(USE_LANGUAGE_CODE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(USE_LANGUAGE_CODE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            } 
            if (iConv.ISNull(W_HALF_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(W_HALF_TYPE_NAME)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(HOMETAX_LOGIN_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(HOMETAX_LOGIN_ID)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            if (iConv.ISNull(WRITE_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", Get_Edit_Prompt(WRITE_DATE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }

            if (IGR_FILE_SUM.RowCount < 1)
            {
                MessageBoxAdv.Show("생성할 원천징수의무자 자료 건수가 존재하지 않습니다. 조회후 다시 실행하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return vRETURN;
            }
            vRETURN = "Y";
            return vRETURN;
        }

        private void Button_Control(string pEnabled_YN)
        {
            if (pEnabled_YN == "Y")
            {
                BTN_CANCEL_CLOSED.Enabled = true;
                BTN_SET_CLOSED.Enabled = true;
                BTN_CREATE.Enabled = true;
                BTN_SET_FILE.Enabled = true; 
            }
            else
            {
                BTN_CANCEL_CLOSED.Enabled = false;
                BTN_SET_CLOSED.Enabled = false;
                BTN_CREATE.Enabled = false;
                BTN_SET_FILE.Enabled = false; 
            }

        }

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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
        }

        #endregion;
         
        #region ----- Text File Export Methods ----

        private void Encrypt_ExportTXT(string pFileName, string pFILE_TYPE, ISDataAdapter pData)
        {
            object vFIX_STRING = null;
            int vCountRow = pData.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            //전산매체 암호화 암호 입력 받기.
            DialogResult vdlgResult;
            object vENCRYPT_PASSWORD = String.Empty;
            HRMF0796_FILE vHRMF0796_FILE = new HRMF0796_FILE(isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0796_FILE.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            {
                vENCRYPT_PASSWORD = vHRMF0796_FILE.Get_Encrypt_Password;
            }

            if (iConv.ISNull(vENCRYPT_PASSWORD) == string.Empty)
            {
                return;
            }

            Button_Control("N");  //버튼 사용 불가 만들기.
            if (pFILE_TYPE == "ADJUST")
            {
                vFIX_STRING = "C";
            }
            else if (pFILE_TYPE == "MEDICAL")
            {
                vFIX_STRING = "CA";
            }
            else if (pFILE_TYPE == "DONATION")
            {
                vFIX_STRING = "H";
            }

            isAppInterfaceAdv1.OnAppMessage("Export Text Start...");

            string vSaveTextFileName = String.Empty;
            string vFileName = string.Empty;
            string vFilePath = "C:\\ersdata";

            int euckrCodepage = 51949;
            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

            //파일 경로 디렉토리 존재 여부 체크(없으면 생성).
            if (System.IO.Directory.Exists(vFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(vFilePath);
            }

            vFileName = String.Format("{0}{1}", vFIX_STRING, iConv.ISNull(pFileName).Replace("-", "").Substring(0, 7));
            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = vFileName;
            saveFileDialog1.DefaultExt = ".txt";  // String.Format(".{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            //System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(vFilePath);
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Text Files (*.txt)|*.txt";//String.Format("Text Files (*.{0})|*.{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                Application.DoEvents();

                vSaveTextFileName = saveFileDialog1.FileName;
                try
                {
                    vWriteFile = System.IO.File.Open(vSaveTextFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    foreach (DataRow cRow in pData.OraSelectData.Rows)
                    {
                        vSaveString = new System.Text.StringBuilder();  //초기화.
                        vSaveString.Append(cRow["REPORT_FILE"]);
                        vSaveString.Append("\r\n");

                        //기존
                        //byte[] vSaveBytes = new System.Text.UnicodeEncoding().GetBytes(vSaveString.ToString());

                        //신규.
                        System.Text.Encoding vEUCKR = System.Text.Encoding.GetEncoding(euckrCodepage);
                        byte[] vSaveBytes = vEUCKR.GetBytes(vSaveString.ToString());

                        int vSaveStrigLength = vSaveBytes.Length;
                        vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);
                    }
                }
                catch (System.Exception ex)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    string vMessage = ex.Message;
                    isAppInterfaceAdv1.OnAppMessage(vMessage);
                    Application.DoEvents();
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                }
                isAppInterfaceAdv1.OnAppMessage("Complete, Export Text~!");
                vWriteFile.Dispose();

                //기존 동일한 파일 삭제.
                if (System.IO.File.Exists(vSaveTextFileName) == false)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    Application.DoEvents();
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    MessageBoxAdv.Show("암호화 대상 전자파일이 존재하지 않습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                nRet = 0;
                inputPath = vSaveTextFileName;// "20120410.201";//pFileName;
                OutputPath = string.Format("{0}.erc", vSaveTextFileName);
                Password = vENCRYPT_PASSWORD.ToString();
                DSFC_OPT_OVERWRITE_OUTPUT = 1;
                nRet = DSFC_EncryptFile(0, inputPath, OutputPath, Password, DSFC_OPT_OVERWRITE_OUTPUT);
                if (nRet != 0)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    Application.DoEvents();
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    MessageBox.Show(String.Format("Encrypt Error : {0}", nRet), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                System.IO.File.Delete(vSaveTextFileName);
                System.IO.File.Copy(OutputPath, inputPath, true);
                System.IO.File.Delete(OutputPath);                
            }
            Button_Control("Y");  //버튼 사용 만들기.
            Application.DoEvents();
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void ExportTXT(string pFileName, string pFILE_TYPE, ISDataAdapter pData)
        {
            object vFIX_STRING = null;
            int vCountRow = pData.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                V_STATUS.PromptText = "";
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
             
            vFIX_STRING = "SF";

            V_STATUS.PromptText = string.Format("eFile 내려받기 중 {0} :: {1}", "", W_YEAR_YYYY.EditValue, W_HALF_TYPE_NAME.EditValue);
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            int euckrCodepage = 51949;
            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();
            
            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = String.Format("{0}{1}", vFIX_STRING, iConv.ISNull(pFileName).Replace("-", "").Substring(0, 7));
            saveFileDialog1.DefaultExt = String.Format(".{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = String.Format("Text Files (*.{0})|*.{0}", iConv.ISNull(pFileName).Replace("-", "").Substring(7, 3));
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                Application.DoEvents();

                string vsSaveTextFileName = saveFileDialog1.FileName;
                try
                {
                    vWriteFile = System.IO.File.Open(vsSaveTextFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    foreach (DataRow cRow in pData.OraSelectData.Rows)
                    {
                        vSaveString = new System.Text.StringBuilder();  //초기화.
                        vSaveString.Append(cRow["REPORT_FILE"]);
                        vSaveString.Append("\r\n");

                        //기존
                        //byte[] vSaveBytes = new System.Text.UnicodeEncoding().GetBytes(vSaveString.ToString());

                        //신규.
                        System.Text.Encoding vEUCKR = System.Text.Encoding.GetEncoding(euckrCodepage);
                        byte[] vSaveBytes = vEUCKR.GetBytes(vSaveString.ToString());

                        int vSaveStrigLength = vSaveBytes.Length;
                        vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);                        
                    }
                }
                catch (System.Exception ex)
                {
                    Button_Control("Y");  //버튼 사용 만들기.
                    string vMessage = ex.Message;
                    isAppInterfaceAdv1.OnAppMessage(vMessage);

                    V_STATUS.PromptText = "";
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents(); 
                }
                vWriteFile.Dispose();
            }
            Button_Control("Y");  //버튼 사용 만들기.
            V_STATUS.PromptText = string.Format("eFile 내려받기 완료 {0} :: {1}", "", W_YEAR_YYYY.EditValue, W_HALF_TYPE_NAME.EditValue);
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        public void ExportTXT_File(ISDataAdapter pData)
        {
        //    int vCountRow = pData.OraSelectData.Rows.Count;
        //    if (vCountRow < 1)
        //    {
        //        return;
        //    }

        //    isAppInterfaceAdv1.OnAppMessage("Export Text Start...");

        //    System.IO.Stream vWrite = null; ;
        //    System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

        //    saveFileDialog1.Title = "Save File";
        //    saveFileDialog1.FileName = WRITE_DATE.DateTimeValue.ToShortDateString().Replace("-", "");
        //    saveFileDialog1.DefaultExt = ".101";
        //    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
        //    saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
        //    saveFileDialog1.Filter = "Text Files (*.101)|*.101";
        //    if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //    {
        //        Application.UseWaitCursor = true;
        //        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //        Application.DoEvents();

        //        string vsSaveTextFileName = saveFileDialog1.FileName;
        //        try
        //        {
        //            //vWriteFile = System.IO.File.Open(vsSaveTextFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
        //            vWrite = System.IO.File.OpenWrite(vsSaveTextFileName);
        //            foreach (DataRow cRow in pData.OraSelectData.Rows)
        //            {
        //                vSaveString = new System.Text.StringBuilder();  //초기화.
        //                vSaveString.Append(cRow["REPORT_FILE"]);
        //                vSaveString.Append("\r\n");

        //                System.IO.StreamWriter(vWrite, Encoding.Default);

        //                //byte[] vSaveBytes = new System.Text.UnicodeEncoding().GetBytes(vSaveString.ToString());
        //                //int vSaveStrigLength = vSaveBytes.Length;
        //                //vWriteFile.Write(vSaveBytes, 0, vSaveStrigLength);
        //            }
        //        }
        //        catch (System.Exception ex)
        //        {
        //            string vMessage = ex.Message;
        //            isAppInterfaceAdv1.OnAppMessage(vMessage);
        //            Application.DoEvents();
        //            Application.UseWaitCursor = false;
        //            this.Cursor = System.Windows.Forms.Cursors.Default;
        //        }

        //        isAppInterfaceAdv1.OnAppMessage("Export Text End");
        //        vWriteFile.Dispose();
        //    }
        //    Application.DoEvents();
        //    Application.UseWaitCursor = false;
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        } 

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_eFILE_INFO.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_eFILE_INFO.IsFocused)
                    {
                        IDA_eFILE_INFO.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_eFILE_INFO.IsFocused)
                    {
                        IDA_eFILE_INFO.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0796_Load(object sender, EventArgs e)
        {
            DefaultCorporation();
        }

        private void HRMF0796_Shown(object sender, EventArgs e)
        {
            

            //기준일자.
            W_YEAR_YYYY.EditValue = iDate.ISYear(DateTime.Today);
            START_DATE.EditValue = iDate.ISMonth_1st(string.Format("{0}-01", iDate.ISYear(DateTime.Today)));
            END_DATE.EditValue = iDate.ISMonth_Last(string.Format("{0}-12", DateTime.Today.Year));

            W_YEAR_YYYY.Focus();
            WRITE_DATE.EditValue = DateTime.Today;
        }

        private void BTN_CREATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(W_YEAR_YYYY.EditValue) == string.Empty)
            {
                W_YEAR_YYYY.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_YEAR_YYYY))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (iConv.ISNull(W_HALF_TYPE.EditValue) == string.Empty)
            {
                W_HALF_TYPE_NAME.Focus();
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_HALF_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10067"), "Qeustion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            Button_Control("N");  //버튼 사용 불가 만들기. 

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            IDC_SET_MAIN.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_MAIN.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_MAIN.GetCommandParamValue("O_MESSAGE"));
            if(vSTATUS == "F")
            {
                Button_Control("Y");
                if (vMESSAGE != String.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (vMESSAGE != String.Empty)
            {
                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Button_Control("Y"); 
        }

        private void IGR_ADJUST_FILE_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }          
            IGR_FILE_SUM.LastConfirmChanges();       
            IDA_FILE_SUM.OraSelectData.AcceptChanges();
            IDA_FILE_SUM.Refillable = true;            
        }

        private void IGR_FILE_PAYMENT_SUM_CellDoubleClick(object pSender)
        {

        }

        private void START_DATE_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            END_DATE.EditValue = string.Format("{0}-12-31", iDate.ISYear(START_DATE.EditValue));
        }


        private void BTN_YEAR_ADJUST_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Button_Control("N");  //버튼 사용 불가 만들기. 

            int vIDX_SELECT_YN = IGR_FILE_SUM.GetColumnToIndex("SELECT_YN");
            int vIDX_OPERATING_UNIT_ID = IGR_FILE_SUM.GetColumnToIndex("OPERATING_UNIT_ID");
            int vIDX_VAT_NUMBER = IGR_FILE_SUM.GetColumnToIndex("VAT_NUMBER");
            int vIDX_TAX_OFFICE_CODE = IGR_FILE_SUM.GetColumnToIndex("TAX_OFFICE_CODE"); 
            for (int r = 0; r < IGR_FILE_SUM.RowCount; r++)
            {
                if (iConv.ISNull(IGR_FILE_SUM.GetCellValue(r, vIDX_SELECT_YN)) == "Y")
                {
                    IGR_FILE_SUM.CurrentCellMoveTo(r, vIDX_SELECT_YN);

                    if (EXPORT_VALIDATE() != "Y")
                    {
                        Button_Control("Y");
                        return;
                    }

                    object vOPERATING_UNIT_ID = IGR_FILE_SUM.GetCellValue(r, vIDX_OPERATING_UNIT_ID);
                    string vVAT_NUMBER = iConv.ISNull(IGR_FILE_SUM.GetCellValue(r, vIDX_VAT_NUMBER));
                    object vTAX_OFFICE_CODE = IGR_FILE_SUM.GetCellValue(r, vIDX_TAX_OFFICE_CODE);
                    if (iConv.ISNull(vOPERATING_UNIT_ID) == string.Empty)
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.UseWaitCursor = false;
                        Application.DoEvents();
                        Button_Control("Y");

                        MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", "사업장 정보"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (vVAT_NUMBER == string.Empty)
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.UseWaitCursor = false;
                        Application.DoEvents();
                        Button_Control("Y");

                        MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", "사업자번호"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (iConv.ISNull(vTAX_OFFICE_CODE) == string.Empty)
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.UseWaitCursor = false;
                        Application.DoEvents();
                        Button_Control("Y");

                        MessageBoxAdv.Show(string.Format("{0}은(는)은 필수입니다. 확인하세요", "관할 세무서"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    //검증.
                    IDC_GET_FILE_CHECK_P.SetCommandParamValue("P_OPERATING_UNIT_ID", vOPERATING_UNIT_ID);
                    IDC_GET_FILE_CHECK_P.ExecuteNonQuery();
                    string vSTATUS = iConv.ISNull(IDC_GET_FILE_CHECK_P.GetCommandParamValue("O_STATUS"));
                    string vMESSAGE = iConv.ISNull(IDC_GET_FILE_CHECK_P.GetCommandParamValue("O_MESSAGE"));
                    if(vSTATUS == "F")
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.UseWaitCursor = false;
                        Application.DoEvents();
                        Button_Control("Y");
                        if (vMESSAGE != String.Empty)
                        {
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    } 

                    //파일 생성//
                    V_STATUS.PromptText = string.Format("eFile 생성중 : {0} :: {1}", "", W_YEAR_YYYY.EditValue, W_HALF_TYPE_NAME.EditValue);
                    Application.DoEvents();
                    IDC_SET_FILE_MAIN.SetCommandParamValue("P_OPERATING_UNIT_ID", vOPERATING_UNIT_ID);
                    IDC_SET_FILE_MAIN.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_FILE_MAIN.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_FILE_MAIN.GetCommandParamValue("O_MESSAGE"));
                    if(vSTATUS == "F")
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.UseWaitCursor = false;
                        Application.DoEvents();
                        Button_Control("Y");
                        if (vMESSAGE != String.Empty)
                        {
                            V_STATUS.PromptText = "";
                            MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return;
                    }

                    V_STATUS.PromptText = string.Format("eFile 내려받기 시작 {0} :: {1}", "", W_YEAR_YYYY.EditValue, W_HALF_TYPE_NAME.EditValue);
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.UseWaitCursor = true;
                    Application.DoEvents(); 

                    IDA_FILE_EXPORT.SetDeleteParamValue("P_OPERATING_UNIT_ID", vOPERATING_UNIT_ID);
                    IDA_FILE_EXPORT.Fill();
                    ExportTXT(vVAT_NUMBER, "ADJUST", IDA_FILE_EXPORT);

                    //생성된 자료수 체크//
                    IDC_GET_FILE_COUNT_P.SetCommandParamValue("P_OPERATING_UNIT_ID", vOPERATING_UNIT_ID);
                    IDC_GET_FILE_COUNT_P.ExecuteNonQuery();
                    object vREC_COUNT = IDC_GET_FILE_COUNT_P.GetCommandParamValue("O_REC_COUNT");
                     
                    IGR_FILE_SUM.SetCellValue(r, vIDX_SELECT_YN, "N");

                    V_STATUS.PromptText = string.Format("eFile 내려받기 완료 {0} :: {1}", "", W_YEAR_YYYY.EditValue, W_HALF_TYPE_NAME.EditValue);
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.UseWaitCursor = false;
                    Application.DoEvents();
                }
            }
            Button_Control("Y");
            IGR_FILE_SUM.LastConfirmChanges();
            IDA_FILE_SUM.OraSelectData.AcceptChanges();
            IDA_FILE_SUM.Refillable = true;

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents(); 
        }
           
        #endregion


        #region ----- Lookup Event -----
        
        private void ILA_SUBMIT_AGENT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetParameter("SUBMIT_AGENT", "Y");
        }

        private void ILA_HALF_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetParameter("HALF_TYPE", "Y");
        }
         
        private void ILA_YYYYMM_FR_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(), 4)));
        }

        private void ILA_YYYYMM_TO_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", W_PERIOD_FR.EditValue);
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(), 4)));
        }

        private void ILA_POST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetParameter("POST", "Y");
        }

        private void ILA_TAX_OFFICE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetParameter("TAX_OFFICE", "Y");
        }

        private void ILA_HALF_TYPE_SelectedRowData(object pSender)
        {
            IDC_GET_PERIOD_P.ExecuteNonQuery();
            W_PERIOD_FR.EditValue = IDC_GET_PERIOD_P.GetCommandParamValue("O_PERIOD_FR");
            W_PERIOD_TO.EditValue = IDC_GET_PERIOD_P.GetCommandParamValue("O_PERIOD_TO");
            START_DATE.EditValue = IDC_GET_PERIOD_P.GetCommandParamValue("O_DATE_FR");
            END_DATE.EditValue = IDC_GET_PERIOD_P.GetCommandParamValue("O_DATE_TO");
        }

        private void ILA_YYYYMM_FR_SelectedRowData(object pSender)
        {
            IDC_GET_PERIOD_DATE_P.ExecuteNonQuery();
            START_DATE.EditValue = IDC_GET_PERIOD_DATE_P.GetCommandParamValue("O_DATE_FR");
        }

        private void ILA_YYYYMM_TO_SelectedRowData(object pSender)
        {
            IDC_GET_PERIOD_DATE_P.ExecuteNonQuery();
            END_DATE.EditValue = IDC_GET_PERIOD_DATE_P.GetCommandParamValue("O_DATE_TO");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_eFILE_INFO_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            } 
        }

        private void IDA_eFILE_INFO_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["CORP_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show("업체정보가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["OPERATING_UNIT_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show("사업장정보가 정확하지 않습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["CORP_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은 필수입니다. 확인하세요", Get_Edit_Prompt(CORP_NAME)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["PRESIDENT_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("{0}은(는) 필수입니다. 확인하세요", Get_Edit_Prompt(PRESIDENT_NAME)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            //if (iConv.ISNull(e.Row["VAT_NUMBER"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(string.Format("{0}은(는) 필수입니다. 확인하세요", Get_Edit_Prompt(VAT_NUMBER)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    CORP_NAME_0.Focus();
            //    return;
            //}
            //if (iConv.ISNull(e.Row["TAX_OFFICE_CODE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(string.Format("{0}은(는) 필수입니다. 확인하세요", Get_Edit_Prompt(TAX_OFFICE_CODE)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    CORP_NAME_0.Focus();
            //    return;
            //}
            //if (iConv.ISNull(e.Row["TAX_OFFICE_NAME"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(string.Format("{0}은(는) 필수입니다. 확인하세요", Get_Edit_Prompt(TAX_OFFICE_NAME)), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    CORP_NAME_0.Focus();
            //    return;
            //}
        }

        private void IDA_ADJUST_FILE_LIST_PreNewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            //if (pBindingManager.DataRow == null)
            //{
            //    return;
            //}
            //int vIDX_SELECT_YN = IGR_ADJUST_FILE_LIST.GetColumnToIndex("SELECT_YN");
            //if (iConv.ISNull(pBindingManager.DataRow["VAT_NUMBER"]) == string.Empty)
            //{
            //    IGR_ADJUST_FILE_LIST.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 0;
            //    IGR_ADJUST_FILE_LIST.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 0;
            //}
            //else
            //{
            //    IGR_ADJUST_FILE_LIST.GridAdvExColElement[vIDX_SELECT_YN].Insertable = 1;
            //    IGR_ADJUST_FILE_LIST.GridAdvExColElement[vIDX_SELECT_YN].Updatable = 1;
            //}
        }

        #endregion
         
    }
}