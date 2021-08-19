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

namespace HRMF0781
{
    public partial class HRMF0781 : Office2007Form
    {
        #region ----- API Dll Import -----

        [DllImport("fcrypt_es.dll")]
        public static extern int DSFC_EncryptFile(int hWnd, string pszPlainFilePathName, string pszEncFilePathName, string pszPassword, uint nOption);        

        #endregion;

        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        string inputPath;
        string OutputPath;
        string Password;
        uint DSFC_OPT_OVERWRITE_OUTPUT;
        int nRet;

        decimal gA04_PERSON_COUNT = 0;
        decimal gA04_PAYMENT = 0;
        decimal gA04_INCOME_TAX = 0;
        decimal gA04_SP_TAX = 0;
        decimal gA04_ADD_TAX = 0;

        #endregion;

        #region ----- Constructor -----

        public HRMF0781()
        {
            InitializeComponent();
        }

        public HRMF0781(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        //업체
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

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (TB_MAIN.SelectedIndex == 0)
            {
                IDA_WITHHOLDING_LIST.Fill();
                IGR_WITHHOLDING_LIST.Focus();
            }
            else if (TB_MAIN.SelectedIndex == 1)
            {
                Search_DB_Detail(WITHHOLDING_DOC_ID.EditValue);
            }
        }

        private void Search_DB_Detail(object pWITHHOLDING_DOC_ID)
        {
            if (iConv.ISNull(pWITHHOLDING_DOC_ID) == string.Empty)
            {
                return;
            }

            IDA_WITHHOLDING_DOC.OraSelectData.AcceptChanges();
            IDA_WITHHOLDING_DOC.Refillable = true;
            
            IDA_WITHHOLDING_DOC.SetSelectParamValue("P_WITHHOLDING_DOC_ID", pWITHHOLDING_DOC_ID);
            IDA_WITHHOLDING_DOC.Fill();
            WITHHOLDING_TYPE_DESC.Focus();

            //연말정산 합계분 값 저장//
            //분납신청,납부금액 등록시 동기화 위함//
            gA04_PERSON_COUNT = iConv.ISDecimaltoZero(A04_PERSON_CNT.EditValue);
            gA04_PAYMENT = iConv.ISDecimaltoZero(A04_PAYMENT_AMT.EditValue);
            gA04_INCOME_TAX = iConv.ISDecimaltoZero(A04_INCOME_TAX_AMT.EditValue);
            gA04_SP_TAX = iConv.ISDecimaltoZero(A04_SP_TAX_AMT.EditValue);
            gA04_ADD_TAX = iConv.ISDecimaltoZero(A04_ADD_TAX_AMT.EditValue);
        }

        private Boolean Check_WITHHOLDING_DOC_Added()
        {
            Boolean Row_Added_Status = false;

            for (int r = 0; r < IDA_WITHHOLDING_DOC.SelectRows.Count; r++)
            {
                if (IDA_WITHHOLDING_DOC.SelectRows[r].RowState == DataRowState.Added)
                {
                    Row_Added_Status = true;
                }
            }
            if (Row_Added_Status == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return (Row_Added_Status);
        }

        private void Printing_DOC(string pOutput_Type)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            HRMF0781_PRINT vHRMF0781_PRINT = new HRMF0781_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vHRMF0781_PRINT.ShowDialog();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (dlgRESULT == DialogResult.OK)
            {
                //인쇄 선택.
                if (vHRMF0781_PRINT.Print_1_YN == "Y")
                {
                    XLPrinting1(pOutput_Type);
                }
                if (vHRMF0781_PRINT.Print_2_YN == "Y")
                {
                    XLPrinting2(pOutput_Type);
                }
                if (vHRMF0781_PRINT.Print_3_YN == "Y")
                {
                    XLPrinting3(pOutput_Type);
                }
                if (vHRMF0781_PRINT.Print_4_YN == "Y")
                {
                    XLPrinting4(pOutput_Type);
                }
                if (vHRMF0781_PRINT.Print_5_YN == "Y")
                {
                    XLPrinting5(pOutput_Type);
                }
                if (vHRMF0781_PRINT.Print_6_YN == "Y")
                {
                    XLPrinting6(pOutput_Type);
                }
            }
            vHRMF0781_PRINT.Dispose();
        }

        #endregion;

        #region ----- Using Dynamic DLL ------

        public class DLLHolder
        {
            //[DllImport("kernel32.dll", EntryPoint = "LoadLibrary")]
            //public static extern int LoadLibrary(string pLibraryName);
            //[DllImport("fcrypt_es.dll")]
            //public static extern int DSFC_EncryptFile(IntPtr hWnd, string pszPlainFilePathName, string pszEncFilePathName, string pszPassword, int nOption);
        }

        #endregion

        #region ----- Export TXT File ------

        private void Export_File()
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            IDA_WITHHOLDING_FILE.Fill();
            if (IDA_WITHHOLDING_FILE.SelectRows.Count < 1)
            {
                isAppInterfaceAdv1.OnAppMessage("Not found Data. Fail, Export file");
                return;
            }

            isAppInterfaceAdv1.OnAppMessage("Export File start...");

            string vSaveFile_name = string.Empty;
            string vFileName = string.Empty;
            string vFileExted = string.Empty;
            string vFilePath  = "C:\\ersdata";

            int euckrCodepage = 51949;
            System.IO.FileStream vWriteFile = null;
            System.Text.StringBuilder vSaveString = new System.Text.StringBuilder();

            //파일명(제출일자).
            vFileName = iDate.ISGetDate(SUBMIT_DATE.EditValue).ToShortDateString(). Replace("-", "");
            
            //신고구분에 따른 확장자 지정.
            if (iConv.ISNull(WITHHOLDING_TYPE.EditValue) == "1")
            {
                vFileExted = "201";
            }
            else if (iConv.ISNull(WITHHOLDING_TYPE.EditValue) == "2")
            {
                vFileExted = "202";
            }
            else if (iConv.ISNull(WITHHOLDING_TYPE.EditValue) == "3")
            {
                vFileExted = "205";
            }
            vFileName = string.Format("{0}.{1}", vFileName, vFileExted);
            
            //파일 경로 디렉토리 존재 여부 체크(없으면 생성).
            if (System.IO.Directory.Exists(vFilePath) == false)
            {
                System.IO.Directory.CreateDirectory(vFilePath);
            }

            saveFileDialog1.Title = "Save File";
            saveFileDialog1.FileName = vFileName;
            saveFileDialog1.DefaultExt = vFileExted;
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(vFilePath);
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = string.Format("Text File(*.{0})|*.{0}", vFileExted);
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                vSaveFile_name = saveFileDialog1.FileName;
                //기존 동일한 파일 삭제.
                if (System.IO.File.Exists(vSaveFile_name) == true)
                {
                    try
                    {
                        System.IO.File.Delete(vSaveFile_name);
                    }
                    catch (Exception ex)
                    {
                        MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                //파일 생성.
                try
                {
                    vWriteFile = System.IO.File.Open(vSaveFile_name, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
                    foreach (DataRow pRow in IDA_WITHHOLDING_FILE.SelectRows)
                    {
                        vSaveString = new System.Text.StringBuilder();
                        vSaveString.Append(pRow["REPORT_FILE"]);
                        vSaveString.Append("\r\n");

                        System.Text.Encoding vEuckr = System.Text.Encoding.GetEncoding(euckrCodepage);
                        byte[] vSavebytes = vEuckr.GetBytes(vSaveString.ToString());

                        int vSaveStringLength = vSavebytes.Length;
                        vWriteFile.Write(vSavebytes, 0, vSaveStringLength);
                    }
                    
                }
                catch(Exception ex)
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                vWriteFile.Dispose();
            }

            Encrypt_File(vSaveFile_name);
            isAppInterfaceAdv1.OnAppMessage("Complete, Export file");
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void Encrypt_File(string pFileName)
        {
            object vEncrypt_PWD;

            IDC_GET_ENCRYPT_PWD.ExecuteNonQuery();
            vEncrypt_PWD = IDC_GET_ENCRYPT_PWD.GetCommandParamValue("O_ENCRYPT_PWD");
            if (iConv.ISNull(vEncrypt_PWD) == string.Empty)
            {
                MessageBoxAdv.Show("전자파일 암호화 비밀번호를 입력하지 않았습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //기존 동일한 파일 삭제.
            if (System.IO.File.Exists(pFileName) ==  false)
            {
                MessageBoxAdv.Show("암호화 대상 전자파일이 존재하지 않습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            nRet = 0;
            inputPath = pFileName;// "20120410.201";//pFileName;
            OutputPath = string.Format("{0}.erc", pFileName);
            Password = vEncrypt_PWD.ToString();
            DSFC_OPT_OVERWRITE_OUTPUT = 1;
            nRet = DSFC_EncryptFile(0, inputPath, OutputPath, Password, DSFC_OPT_OVERWRITE_OUTPUT);
            if(nRet != 0)
            {
                MessageBox.Show("Encrypt Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            System.IO.File.Delete(pFileName);
            System.IO.File.Copy(OutputPath, inputPath, true);
            System.IO.File.Delete(OutputPath);
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
            try
            {
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
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;

        #region ----- XL Print 1 Methods ----

        private void XLPrinting1(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_DOC.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_DOC.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_doc_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                if (iDate.ISGetDate(string.Format("{0}-01", STD_YYYYMM.EditValue)) < iDate.ISGetDate(string.Format("{0}-01", "2016-02")))
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0781_001.xls";
                    //-------------------------------------------------------------------------------------
                }
                else if (iDate.ISGetDate(string.Format("{0}-01", STD_YYYYMM.EditValue)) < iDate.ISGetDate(string.Format("{0}-01", "2018-01")))
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0781_011.xls";
                    //-------------------------------------------------------------------------------------
                }
                else
                {
                    //2016-02년도 변경분//
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0781_018.xlsx";
                    //-------------------------------------------------------------------------------------
                }
                
                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    if (iDate.ISGetDate(string.Format("{0}-01", STD_YYYYMM.EditValue)) < iDate.ISGetDate(string.Format("{0}-01", "2016-02")))
                    {
                        vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_WITHHOLDING_DOC);
                    }
                    else
                    {
                        vPageNumber = xlPrinting.ExcelWrite_11(IDA_PRINT_WITHHOLDING_DOC);
                    }

                    //부표 체크.
                    IDC_GET_WITHHOLDING_DOC_SUB_P.ExecuteNonQuery();
                    string vDOC_SUB_FLAG = iConv.ISNull(IDC_GET_WITHHOLDING_DOC_SUB_P.GetCommandParamValue("O_DOC_SUB_FLAG"));
                    if (vDOC_SUB_FLAG == "Y")
                    {
                        IDA_PRINT_WITHHOLDING_DOC_SUB_01.Fill();
                        IDA_PRINT_WITHHOLDING_DOC_SUB_02.Fill();

                        vPageNumber = xlPrinting.ExcelWrite_11_SUB(IDA_PRINT_WITHHOLDING_DOC_SUB_01, IDA_PRINT_WITHHOLDING_DOC_SUB_02);
                    }

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 2 Methods ----

        private void XLPrinting2(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_2.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_2.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_2_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0781_002.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_WITHHOLDING_2);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 3 Methods ----

        private void XLPrinting3(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_3.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_3.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_3_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0781_003.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_WITHHOLDING_3);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 4 Methods ----

        private void XLPrinting4(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_4.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_4.OraSelectData.Rows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_4_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0781_004.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_WITHHOLDING_4);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 5 Methods ----

        private void XLPrinting5(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_5.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_5.CurrentRows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_5_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0781_005.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_WITHHOLDING_5);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- XL Print 6 Methods ----

        private void XLPrinting6(string pOutput_Type)
        {
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;

            // 데이터 조회.
            IDA_PRINT_WITHHOLDING_6.Fill();
            vCountRow = IDA_PRINT_WITHHOLDING_6.CurrentRows.Count;

            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Withholding_6_";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xls";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0781_006.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite2(IDA_PRINT_WITHHOLDING_6);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

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
                    if (IDA_WITHHOLDING_LIST.IsFocused || IDA_WITHHOLDING_DOC.IsFocused)
                    {
                        if (Check_WITHHOLDING_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_WITHHOLDING_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_WITHHOLDING_DOC.AddOver();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;

                        MONTHLY_YN.CheckBoxValue = "Y";
                        WITHHOLDING_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_WITHHOLDING_LIST.IsFocused || IDA_WITHHOLDING_DOC.IsFocused)
                    {
                        if (Check_WITHHOLDING_DOC_Added() == true)
                        {
                            // INSERT 중인 작업이 존재함.
                            return;
                        }

                        if (IDA_WITHHOLDING_LIST.IsFocused)
                        {
                            TB_MAIN.SelectedIndex = 1;
                            TB_MAIN.SelectedTab.Focus();
                        }

                        IDA_WITHHOLDING_DOC.AddUnder();

                        STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        PAY_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
                        SUBMIT_DATE.EditValue = DateTime.Today;

                        MONTHLY_YN.CheckBoxValue = "Y";
                        WITHHOLDING_TYPE_DESC.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_WITHHOLDING_LIST.IsFocused)
                    {
                        IDA_WITHHOLDING_LIST.Update();
                    }
                    else if (IDA_WITHHOLDING_DOC.IsFocused)
                    {
                        WITHHOLDING_NO.Focus();
                        IDA_WITHHOLDING_DOC.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_WITHHOLDING_LIST.IsFocused)
                    {
                        IDA_WITHHOLDING_LIST.Cancel();
                    }
                    else if (IDA_WITHHOLDING_DOC.IsFocused)
                    {
                        IDA_WITHHOLDING_DOC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_WITHHOLDING_LIST.IsFocused)
                    {
                        IDA_WITHHOLDING_LIST.Delete();
                    }
                    else if (IDA_WITHHOLDING_DOC.IsFocused)
                    {
                        IDA_WITHHOLDING_DOC.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    Printing_DOC("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    Printing_DOC("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0781_Load(object sender, EventArgs e)
        {
            IDA_WITHHOLDING_DOC.FillSchema();
        }

        private void HRMF0781_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            SUBMIT_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }

        private void IGR_WITHHOLDING_LIST_CellDoubleClick(object pSender)
        {
            if (IGR_WITHHOLDING_LIST.Row < 1)
            {
                return;
            }
            TB_MAIN.SelectedIndex = 1;
            TB_MAIN.SelectedTab.Focus();

            Search_DB_Detail(IGR_WITHHOLDING_LIST.GetCellValue("WITHHOLDING_DOC_ID"));
        }

        private void BTN_PROCESS_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //update.
            IDA_WITHHOLDING_DOC.Update();

            if (iConv.ISNull(WITHHOLDING_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WITHHOLDING_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 귀속년월 및 지급년월이 매년도 2월일 경우 집계데이터 선택 화면 표시.
            object vSET_TYPE = "3";  // 집계데이터 선택 : 1-연말정산데이터, 2-매월징수분 + 연말정산데이터, 3-매월징수분.
            if (iConv.ISNull(STD_YYYYMM.EditValue).Substring(5, 2) == "02" && iConv.ISNull(PAY_YYYYMM.EditValue).Substring(5, 2) == "02")
            {
                DialogResult dlgRESULT;
                HRMF0781_SET vHRMF0781 = new HRMF0781_SET(isAppInterfaceAdv1.AppInterface);
                dlgRESULT =vHRMF0781.ShowDialog();
                if (dlgRESULT == DialogResult.Cancel)
                {
                    return;
                }
                vSET_TYPE = vHRMF0781.Set_Type;
                vHRMF0781.Dispose();
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_MAIN_WITHHOLDING.SetCommandParamValue("P_SET_TYPE", vSET_TYPE);
            IDC_MAIN_WITHHOLDING.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_MAIN_WITHHOLDING.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_MAIN_WITHHOLDING.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_MAIN_WITHHOLDING.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
         
            // requery.
            Search_DB_Detail(WITHHOLDING_DOC_ID.EditValue);
        }

        private void BTN_CLOSED_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(WITHHOLDING_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WITHHOLDING_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;
            IDC_CLOSED_WITHHOLDING.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_WITHHOLDING.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_WITHHOLDING.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_WITHHOLDING.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BTN_CLOSED_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(WITHHOLDING_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WITHHOLDING_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = null;

            IDC_CLOSED_CANCEL_WITHHOLDING.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_CLOSED_CANCEL_WITHHOLDING.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_CLOSED_CANCEL_WITHHOLDING.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (IDC_CLOSED_CANCEL_WITHHOLDING.ExcuteError || vSTATUS == "F")
            {
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BTN_FILE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(WITHHOLDING_DOC_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WITHHOLDING_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgResult;
            HRMF0781_FILE vHRMF0781_FILE = new HRMF0781_FILE(isAppInterfaceAdv1.AppInterface, CORP_ID_0.EditValue);
            dlgResult = vHRMF0781_FILE.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                //전산매체 작성.
                Export_File();
            }
            vHRMF0781_FILE.Dispose();
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_SUB_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iConv.ISNull(WITHHOLDING_DOC_ID.EditValue) == string.Empty)
            {
                return;
            }

            HRMF0781_SUB vHRMF0781_SUB = new HRMF0781_SUB(this.MdiParent, isAppInterfaceAdv1.AppInterface, WITHHOLDING_DOC_ID.EditValue);
            DialogResult vdlgResult = vHRMF0781_SUB.ShowDialog();
            if(vdlgResult == DialogResult.OK)
            {
                Search_DB_Detail(WITHHOLDING_DOC_ID.EditValue);
            }
            vHRMF0781_SUB.Dispose();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_CALENDAR_YEAR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CALENDAR_YEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today, 1));
        }

        private void ILA_STD_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_PAY_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ILA_WITHHOLDING_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "WITHHOLDING_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter event ------
        
        private void IDA_WITHHOLDING_DOC_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["WITHHOLDING_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(WITHHOLDING_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["STD_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(STD_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["SUBMIT_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(SUBMIT_DATE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISDecimaltoZero(e.Row["NEXT_REFUND_TAX_AMT"]) < 0)
            {
                MessageBoxAdv.Show("(20).차월이월환급세액은 (-)일 수 없습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISDecimaltoZero(e.Row["NEXT_REFUND_TAX_AMT"]) < iConv.ISDecimaltoZero(e.Row["REQUEST_REFUND_TAX_AMT"]))
            {
                MessageBoxAdv.Show("(21).환급신청액이 (20)차월이월환급세액보다 많습니다. 확인하세요", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (CHECK_A10_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A20_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A30_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A40_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A47_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A50_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A60_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A69_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A70_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A80_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
            if (CHECK_A90_PAY_TAX_AMT() == false)
            {
                e.Cancel = true;
                return;
            }
        }
        
        #endregion

        #region ----- 합계 금액 동기화 -----

        private void A04_PERSON_CNT_EditValueChanged(object pSender)
        {
            //분납신청,납부금액 등록시 동기화 위함//
            gA04_PERSON_COUNT = iConv.ISDecimaltoZero(A04_PERSON_CNT.EditValue);
        }

        private void A04_INCOME_TAX_AMT_EditValueChanged(object pSender)
        {

        }

        private void A04_SP_TAX_AMT_EditValueChanged(object pSender)
        {
            //분납신청,납부금액 등록시 동기화 위함// 
            gA04_SP_TAX = iConv.ISDecimaltoZero(A04_SP_TAX_AMT.EditValue);
        }

        private void A04_ADD_TAX_AMT_EditValueChanged(object pSender)
        {

        }

        private void A05_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A05_PERSON_CNT();
        }

        private void A05_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A05_INCOME_TAX_AMT();
        }

        private void A05_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A05_SP_TAX_AMT();
        }

        private void A05_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A05_ADD_TAX_AMT();
        }
         
        private void A10_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A10_PERSON_CNT();
        }

        private void A10_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A10_PAYMENT_AMT();            
        }

        private void A10_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A10_INCOME_TAX_AMT();
        }

        private void A10_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {            
            SUM_A10_SP_TAX_AMT();
        }

        private void A10_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A10_ADD_TAX_AMT();
        }

        private void A10_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A10_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A10_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private void A10_PAY_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_SP_TAX_AMT();
        }


        private void A20_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A20_PERSON_CNT();          
        }

        private void A20_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A20_PAYMENT_AMT(); 
        }

        private void A20_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A20_INCOME_TAX_AMT(); 
        }

        private void A20_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A20_ADD_TAX_AMT(); 
        }

        private void A20_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A20_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A20_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A30_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A30_PERSON_CNT(); 
        }

        private void A30_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A30_PAYMENT_AMT(); 
        }

        private void A30_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A30_INCOME_TAX_AMT(); 
        }

        private void A30_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A30_SP_TAX_AMT();
        }

        private void A30_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A30_ADD_TAX_AMT();
        }

        private void A30_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A30_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A30_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private void A30_PAY_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_SP_TAX_AMT();
        }


        private void A40_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A40_PERSON_CNT();
        }

        private void A40_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A40_PAYMENT_AMT();
        }

        private void A40_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A40_INCOME_TAX_AMT();
        }

        private void A40_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A40_ADD_TAX_AMT();
        }

        private void A40_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A40_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A40_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A47_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A47_PERSON_CNT();
        }

        private void A47_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A47_PAYMENT_AMT();
        }

        private void A47_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A47_INCOME_TAX_AMT();
        }

        private void A47_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A47_ADD_TAX_AMT();
        }

        private void A47_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A47_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A47_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A50_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A50_PERSON_CNT();
        }

        private void A50_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A50_PAYMENT_AMT();
        }

        private void A50_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A50_INCOME_TAX_AMT();
        }

        private void A50_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A50_SP_TAX_AMT();
        }

        private void A50_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A50_ADD_TAX_AMT();
        }

        private void A50_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A50_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A50_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private void A50_PAY_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_SP_TAX_AMT();
        }


        private void A60_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A60_PERSON_CNT();
        }

        private void A60_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A60_PAYMENT_AMT();
        }

        private void A60_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A60_INCOME_TAX_AMT();
        }

        private void A60_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A60_SP_TAX_AMT();
        }

        private void A60_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A60_ADD_TAX_AMT();
        }

        private void A60_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A60_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A60_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private void A60_PAY_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_SP_TAX_AMT();
        }


        private void A69_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A69_PERSON_CNT();
        }

        private void A69_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A69_INCOME_TAX_AMT();
        }

        private void A69_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A69_ADD_TAX_AMT();
        }

        private void A69_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A69_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A69_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A70_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A70_PERSON_CNT();
        }

        private void A70_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A70_PAYMENT_AMT(); 
        }

        private void A70_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A70_INCOME_TAX_AMT(); 
        }

        private void A70_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A70_ADD_TAX_AMT(); 
        }

        private void A70_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A70_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A70_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A80_PERSON_CNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A80_PERSON_CNT();  
        }

        private void A80_PAYMENT_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A80_PAYMENT_AMT();  
        }

        private void A80_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A80_INCOME_TAX_AMT();  
        }

        private void A80_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A80_ADD_TAX_AMT(); 
        }

        private void A80_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A80_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }  
        }

        private void A80_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }


        private void A90_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A90_INCOME_TAX_AMT(); 
        }

        private void A90_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A90_SP_TAX_AMT(); 
        }

        private void A90_ADD_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A90_ADD_TAX_AMT();  
        }

        private void A90_THIS_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (SUM_A90_THIS_REFUND_TAX_AMT() == false)
            {
                e.Cancel = true;
            }
        }

        private void A90_PAY_INCOME_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private void A90_PAY_SP_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            TOTAL_PAY_SP_TAX_AMT();
        }


        private void RECEIVE_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_REFUND_BALANCE_AMT(); 
        }

        private void ALREADY_REFUND_TAX_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_REFUND_BALANCE_AMT();           
        }

        private void FINANCIAL_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_ADJUST_REFUND_TAX_AMT(); 
        }

        private void ETC_REFUND_FINANCIAL_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_ADJUST_REFUND_TAX_AMT(); 
        }

        private void ETC_REFUND_MERGER_AMT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            CAL_ADJUST_REFUND_TAX_AMT(); 
        }

        #endregion

        #region ----- Sum or Total -----


        private void SUM_A05_PERSON_CNT()
        {
            A04_PERSON_CNT.EditValue =iConv.ISDecimaltoZero(A05_PERSON_CNT.EditValue);

            SUM_A10_PERSON_CNT(); 
        }
         
        private void SUM_A05_INCOME_TAX_AMT()
        {
            A04_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A05_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A06_INCOME_TAX_AMT.EditValue);

            SUM_A10_INCOME_TAX_AMT();
        }

        private void SUM_A05_SP_TAX_AMT()
        {
            A04_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A05_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_SP_TAX_AMT.EditValue);

            SUM_A10_SP_TAX_AMT();
        }

        private void SUM_A05_ADD_TAX_AMT()
        {
            A04_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A05_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_ADD_TAX_AMT.EditValue);

            SUM_A10_ADD_TAX_AMT();
        }

        private void SUM_A10_PERSON_CNT()
        {
            A10_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A01_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A03_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A04_PERSON_CNT.EditValue);

            TOTAL_PERSON_CNT();
        }

        private void SUM_A10_PAYMENT_AMT()
        {            
            A10_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A01_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A03_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A04_PAYMENT_AMT.EditValue);

            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A10_INCOME_TAX_AMT()
        {
            A10_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A01_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A02_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A03_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A06_INCOME_TAX_AMT.EditValue);
                                            //iConv.ISDecimaltoZero(A04_INCOME_TAX_AMT.EditValue);
            A10_PAY_TAX_AMT();                  //납부세액
            CAL_A10_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산       
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A10_SP_TAX_AMT()
        {
            A10_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A01_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_SP_TAX_AMT.EditValue);
                                        //iConv.ISDecimaltoZero(A04_SP_TAX_AMT.EditValue);

            A10_PAY_TAX_AMT();                  //납부세액
            CAL_A10_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산       
            TOTAL_SP_TAX_AMT();
        }

        private void SUM_A10_ADD_TAX_AMT()
        {
            A10_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A01_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A02_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A03_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A06_ADD_TAX_AMT.EditValue);
                                        //iConv.ISDecimaltoZero(A04_ADD_TAX_AMT.EditValue);


            A10_PAY_TAX_AMT();                  //납부세액
            CAL_A10_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산          
            TOTAL_ADD_TAX_AMT();
        }

        private bool SUM_A10_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A10_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A10_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A10_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            A10_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                A10_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;            
        }

        private bool CHECK_A10_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세            
            if (iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있는경우
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT + 
                                        iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue) + 
                iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A10_PAY_TAX_AMT()
        {            
            //A10. 근로소득 가감계-당월 조정 환급세액 및 납부세액 
            A10_PAY_INCOME_TAX_AMT.EditValue = 0;
            A10_PAY_SP_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue) > 0)
            {
                A10_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue) > 0)
            {
                A10_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세 
            if (iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있을경우
                A10_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급             
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A10_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue);
            decimal vSP_TAX_AMT = iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue);

            //초기화.
            A10_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             

            //(9)당월 조정환금세액 
            //소득세
            if (vINCOME_TAX_AMT >= 0)
            {
                if (NEXT_REFUND_TAX_AMT_F() <= 0)
                {
                    A10_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                }
                else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
                {
                    A10_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                    A10_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                        iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue);
                }
                else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
                {
                    A10_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                    A10_PAY_INCOME_TAX_AMT.EditValue = 0;
                }
            }
            //농특세
            if (vSP_TAX_AMT >= 0)
            {
                if (NEXT_REFUND_TAX_AMT_F() <= 0)
                {
                    A10_PAY_SP_TAX_AMT.EditValue = vSP_TAX_AMT;
                }
                else if (vSP_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
                {
                    A10_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue) +
                                                        NEXT_REFUND_TAX_AMT_F();
                    A10_PAY_SP_TAX_AMT.EditValue = (vINCOME_TAX_AMT + vSP_TAX_AMT) -
                                                    iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue);
                }
                else if (vSP_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
                {
                    A10_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue) +
                                                        vSP_TAX_AMT;
                    A10_PAY_SP_TAX_AMT.EditValue = 0;
                }
            }
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
        }


        private void SUM_A20_PERSON_CNT()
        {
            A20_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A21_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A22_PERSON_CNT.EditValue);

            TOTAL_PERSON_CNT();
        }

        private void SUM_A20_PAYMENT_AMT()
        {
            A20_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A21_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A22_PAYMENT_AMT.EditValue);

            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A20_INCOME_TAX_AMT()
        {
            A20_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A21_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A22_INCOME_TAX_AMT.EditValue);

            A20_PAY_TAX_AMT();                  //납부세액
            CAL_A20_THIS_REFUND_TAX_AMT();           //당월 조정환급세액 계산     
            TOTAL_INCOME_TAX_AMT(); 
        }

        private void SUM_A20_ADD_TAX_AMT()
        {
            A20_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A21_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A22_ADD_TAX_AMT.EditValue);
            A20_PAY_TAX_AMT();                  //납부세액
            CAL_A20_THIS_REFUND_TAX_AMT();          //당월 조정환급세액 계산       
            TOTAL_ADD_TAX_AMT();
        }
         
        private bool SUM_A20_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A20_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A20_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A20_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A20_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A20_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A20_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue);
            }
            
            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A20_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A20_PAY_TAX_AMT()
        {
            //A20. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A20_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue) > 0)
            {
                A20_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue) > 0)
            {
                A20_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue);
            }

            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A20_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue); 
            
            //초기화.
            A20_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A20_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A20_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A20_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A20_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A20_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A20_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }

                
        private void SUM_A30_PERSON_CNT()
        {
            A30_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A25_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A26_PERSON_CNT.EditValue);

            TOTAL_PERSON_CNT();
        }

        private void SUM_A30_PAYMENT_AMT()
        {
            A30_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A25_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A26_PAYMENT_AMT.EditValue);

            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A30_INCOME_TAX_AMT()
        {
            A30_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A25_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A26_INCOME_TAX_AMT.EditValue);
            A30_PAY_TAX_AMT();                  //납부세액
            CAL_A30_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A30_SP_TAX_AMT()
        {
            A30_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A26_SP_TAX_AMT.EditValue);
            A30_PAY_TAX_AMT();                  //납부세액
            CAL_A30_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_SP_TAX_AMT();
        }

        private void SUM_A30_ADD_TAX_AMT()
        {
            A30_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A25_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A26_ADD_TAX_AMT.EditValue);

            A30_PAY_TAX_AMT();                  //납부세액
            CAL_A30_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_ADD_TAX_AMT();
        }

        private bool SUM_A30_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A30_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A30_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A30_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            A30_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                A30_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A30_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세            
            if (iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있는경우
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A30_PAY_TAX_AMT()
        {
            //A30. 사업소득 가감계-당월 조정 환급세액 및 납부세액 
            A30_PAY_INCOME_TAX_AMT.EditValue = 0;
            A30_PAY_SP_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            {
                A30_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) > 0)
            {
                A30_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세 
            if (iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있을경우
                A30_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급         
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A30_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            decimal vSP_TAX_AMT = iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);

            //초기화.
            A30_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A30_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A30_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A30_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A30_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A30_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            //농특세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A30_PAY_SP_TAX_AMT.EditValue = vSP_TAX_AMT;
            }
            else if (vSP_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A30_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
                                                    NEXT_REFUND_TAX_AMT_F();
                A30_PAY_SP_TAX_AMT.EditValue = (vINCOME_TAX_AMT + vSP_TAX_AMT) -
                                                iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vSP_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A30_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
                                                    vSP_TAX_AMT;
                A30_PAY_SP_TAX_AMT.EditValue = 0;
            }

            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }

        
        private void SUM_A40_PERSON_CNT()
        {
            A40_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A41_PERSON_CNT.EditValue) +
                                                       iConv.ISDecimaltoZero(A42_PERSON_CNT.EditValue) +
                                                       iConv.ISDecimaltoZero(A43_PERSON_CNT.EditValue) +
                                                       iConv.ISDecimaltoZero(A44_PERSON_CNT.EditValue); 

            TOTAL_PERSON_CNT();
        }

        private void SUM_A40_PAYMENT_AMT()
        {
            A40_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A41_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A42_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A43_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A44_PAYMENT_AMT.EditValue) ;

            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A40_INCOME_TAX_AMT()
        {
            A40_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A41_INCOME_TAX_AMT.EditValue) +
                                                              iConv.ISDecimaltoZero(A42_INCOME_TAX_AMT.EditValue) +
                                                              iConv.ISDecimaltoZero(A43_INCOME_TAX_AMT.EditValue) +
                                                               iConv.ISDecimaltoZero(A44_INCOME_TAX_AMT.EditValue) ;

            A40_PAY_TAX_AMT();                  //납부세액
            CAL_A40_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A40_ADD_TAX_AMT()
        {
            A40_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A41_ADD_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A42_ADD_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A43_ADD_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A44_ADD_TAX_AMT.EditValue);

            A40_PAY_TAX_AMT();                  //납부세액
            CAL_A40_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_ADD_TAX_AMT();
        }

        private void SUM_A40_PAY_INCOME_TAX_AMT_TOT()
        {
            A40_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A41_PAY_INCOME_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A42_PAY_INCOME_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A43_PAY_INCOME_TAX_AMT.EditValue) +
                                                         iConv.ISDecimaltoZero(A44_PAY_INCOME_TAX_AMT.EditValue);

           // A40_PAY_TAX_AMT();                  //납부세액
            CAL_A40_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_PAY_INCOME_TAX_AMT();
        }

        private bool SUM_A40_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A40_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A40_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A40_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A40_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A40_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A40_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A40_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A40_PAY_TAX_AMT()
        {
            //A40. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A40_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue) > 0)
            {
                A40_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue) > 0)
            {
                A40_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급  
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A40_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue); 

            //초기화.
            A40_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A40_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A40_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A40_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT  -
                                                    iConv.ISDecimaltoZero(A40_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A40_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A40_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            ////합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A47_PERSON_CNT()
        {
            A47_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A48_PERSON_CNT.EditValue) + 
                                        iConv.ISDecimaltoZero(A45_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A46_PERSON_CNT.EditValue);

            TOTAL_PERSON_CNT();
        }

        private void SUM_A47_PAYMENT_AMT()
        {
            A47_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A48_PAYMENT_AMT.EditValue) + 
                                        iConv.ISDecimaltoZero(A45_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A46_PAYMENT_AMT.EditValue);

            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A47_INCOME_TAX_AMT()
        {
            A47_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A48_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A45_INCOME_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A46_INCOME_TAX_AMT.EditValue);

            A47_PAY_TAX_AMT();                  //납부세액
            CAL_A47_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A47_ADD_TAX_AMT()
        {
            A47_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A48_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A45_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A46_ADD_TAX_AMT.EditValue);

            A47_PAY_TAX_AMT();                  //납부세액
            CAL_A47_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산    
            TOTAL_ADD_TAX_AMT();
        }
        
        private bool SUM_A47_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A47_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A47_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A47_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A47_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A47_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A47_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A47_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A47_PAY_TAX_AMT()
        {
            //A47. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A47_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue) > 0)
            {
                A47_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue) > 0)
            {
                A47_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급        
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A47_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue); 

            //초기화.
            A47_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A47_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A47_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A47_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A47_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A47_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A47_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A50_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A50_PAYMENT_AMT()
        {
            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A50_INCOME_TAX_AMT()
        {
            A50_PAY_TAX_AMT();                  //납부세액
            CAL_A50_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산     
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A50_SP_TAX_AMT()
        {
            A50_PAY_TAX_AMT();                  //납부세액
            CAL_A50_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_SP_TAX_AMT();
        }

        private void SUM_A50_ADD_TAX_AMT()
        {
            A50_PAY_TAX_AMT();                  //납부세액
            CAL_A50_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_ADD_TAX_AMT();
        }
         
        private bool SUM_A50_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A50_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A50_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A50_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            A50_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                A50_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A50_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세            
            if (iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있는경우
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A50_PAY_TAX_AMT()
        {
            //A50. 사업소득 가감계-당월 조정 환급세액 및 납부세액 
            A50_PAY_INCOME_TAX_AMT.EditValue = 0;
            A50_PAY_SP_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue) > 0)
            {
                A50_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue) > 0)
            {
                A50_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세 
            if (iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있을경우
                A50_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급             
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A50_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue);
            decimal vSP_TAX_AMT = iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue);

            //초기화.
            A50_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             

            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A50_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A50_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A50_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A50_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A50_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            //농특세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A50_PAY_SP_TAX_AMT.EditValue = vSP_TAX_AMT;
            }
            else if (vSP_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A50_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue) +
                                                    NEXT_REFUND_TAX_AMT_F();
                A50_PAY_SP_TAX_AMT.EditValue = (vINCOME_TAX_AMT + vSP_TAX_AMT) -
                                                iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vSP_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A50_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue) +
                                                    vSP_TAX_AMT;
                A50_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A60_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A60_PAYMENT_AMT()
        {
            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A60_INCOME_TAX_AMT()
        {
            A60_PAY_TAX_AMT();                  //납부세액
            CAL_A60_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산     
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A60_SP_TAX_AMT()
        {
            A60_PAY_TAX_AMT();                  //납부세액
            CAL_A60_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산     
            TOTAL_SP_TAX_AMT();
        }

        private void SUM_A60_ADD_TAX_AMT()
        {
            A60_PAY_TAX_AMT();                  //납부세액
            CAL_A60_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산     
            TOTAL_ADD_TAX_AMT();
        }
 
        private bool SUM_A60_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A60_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A60_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A60_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            A60_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                A60_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A60_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세            
            if (iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있는경우
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A60_PAY_TAX_AMT()
        {
            //A60. 사업소득 가감계-당월 조정 환급세액 및 납부세액 
            A60_PAY_INCOME_TAX_AMT.EditValue = 0;
            A60_PAY_SP_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue) > 0)
            {
                A60_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue) > 0)
            {
                A60_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세 
            if (iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있을경우
                A60_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급             
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A60_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue);
            decimal vSP_TAX_AMT = iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue);

            //초기화.
            A60_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A60_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A60_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A60_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A60_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A60_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            //농특세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A60_PAY_SP_TAX_AMT.EditValue = vSP_TAX_AMT;
            }
            else if (vSP_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A60_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue) +
                                                    NEXT_REFUND_TAX_AMT_F();
                A60_PAY_SP_TAX_AMT.EditValue = (vINCOME_TAX_AMT + vSP_TAX_AMT) -
                                                iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vSP_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A60_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue) +
                                                    vSP_TAX_AMT;
                A60_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A69_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A69_INCOME_TAX_AMT()
        {
            A69_PAY_TAX_AMT();                  //납부세액
            CAL_A69_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A69_ADD_TAX_AMT()
        {
            A69_PAY_TAX_AMT();                  //납부세액
            CAL_A69_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_ADD_TAX_AMT();
        }
 
        private bool SUM_A69_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A69_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A69_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A69_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A69_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A69_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A69_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A69_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A69_PAY_TAX_AMT()
        {
            //A69. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A69_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue) > 0)
            {
                A69_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue) > 0)
            {
                A69_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급         
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A69_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue); 

            //초기화.
            A69_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A69_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A69_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A69_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A69_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A69_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A69_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A70_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A70_PAYMENT_AMT()
        {
            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A70_INCOME_TAX_AMT()
        {
            A70_PAY_TAX_AMT();                  //납부세액
            CAL_A70_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A70_ADD_TAX_AMT()
        {
            A70_PAY_TAX_AMT();                  //납부세액
            CAL_A70_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_ADD_TAX_AMT();
        } 

        private bool SUM_A70_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A70_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A70_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A70_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A70_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A70_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A70_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A70_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A70_PAY_TAX_AMT()
        {
            //A70. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A70_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue) > 0)
            {
                A70_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue) > 0)
            {
                A70_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급        
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A70_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue); 
            //초기화.
            A70_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A70_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A70_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A70_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A70_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A70_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A70_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A80_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A80_PAYMENT_AMT()
        {
            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A80_INCOME_TAX_AMT()
        {
            A80_PAY_TAX_AMT();                  //납부세액
            CAL_A80_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A80_ADD_TAX_AMT()
        {
            A80_PAY_TAX_AMT();                  //납부세액
            CAL_A80_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_ADD_TAX_AMT();
        }
         
        private bool SUM_A80_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A80_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A80_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A80_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A80_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A80_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A80_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A80_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A80_PAY_TAX_AMT()
        {
            //A80. 퇴직소득 가감계-당월 조정 환급세액 및 납부세액 
            A80_PAY_INCOME_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue) > 0)
            {
                A80_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue) > 0)
            {
                A80_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급            
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A80_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue); 

            //초기화.
            A80_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             

            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A80_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A80_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A80_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A80_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A80_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A80_PAY_INCOME_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void SUM_A90_PERSON_CNT()
        {
            TOTAL_PERSON_CNT();
        }

        private void SUM_A90_PAYMENT_AMT()
        {
            TOTAL_PAYMENT_AMT();
        }

        private void SUM_A90_INCOME_TAX_AMT()
        {
            A90_PAY_TAX_AMT();                  //납부세액
            CAL_A90_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_INCOME_TAX_AMT();
        }

        private void SUM_A90_SP_TAX_AMT()
        {
            A90_PAY_TAX_AMT();                  //납부세액
            CAL_A90_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_SP_TAX_AMT();
        }

        private void SUM_A90_ADD_TAX_AMT()
        {
            A90_PAY_TAX_AMT();                  //납부세액
            CAL_A90_THIS_REFUND_TAX_AMT();      //당월 조정환급세액 계산   
            TOTAL_ADD_TAX_AMT();
        }
        
        private void A40_PAY_INCOME_TAX_AMT_TOT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            SUM_A40_PAY_INCOME_TAX_AMT_TOT();

        }
         
        private bool SUM_A90_THIS_REFUND_TAX_AMT()
        {
            //당월 조정 환급세액 합계 
            TOTAL_THIS_REFUND_TAX_AMT();

            //입력금액이 조정대상환급세액보다 크면 오류//
            if (iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) < NEXT_REFUND_TAX_AMT_F())
            {
                MessageBoxAdv.Show("(18)조정대상환급세액보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //납부세액 
            A90_PAY_TAX_AMT();

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue)) < iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue))
            {
                MessageBoxAdv.Show("납부세액((10)소득세등 + (11)농어촌특별세)보다 (9)당월조정환급세액이 많습니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            A90_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) -
                                                iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) < 0)
            {
                A90_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            A90_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue);
            if (iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                A90_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계 
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            return true;
        }

        private bool CHECK_A90_PAY_TAX_AMT()
        {
            //납부세액 검증 
            decimal vTOTOAL_PAY_TAX_AMT = 0;

            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue) > 0)
            {
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세            
            if (iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있는경우
                vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
                                        iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue);
            }

            //납부세액보다 당월조정 환급세액이 많음
            if ((iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) +
                iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            {
                MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // 납부세액((10-소득세등, 11-농어촌특별세) 자동 계산.
        private void A90_PAY_TAX_AMT()
        {
            //A90. 사업소득 가감계-당월 조정 환급세액 및 납부세액 
            A90_PAY_INCOME_TAX_AMT.EditValue = 0;
            A90_PAY_SP_TAX_AMT.EditValue = 0;

            //납부세액-소득세등(가산세 포함)             
            if (iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue) > 0)
            {
                A90_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue) > 0)
            {
                A90_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue) +
                                                    iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue);
            }
            //납부세액-농특세 
            if (iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue) > 0)
            {//납부할 세액이 있을경우
                A90_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue);
            }

            //합계
            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
            CAL_GENERAL_REFUND_AMT();       //(15)일반환급             
        }

        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_A90_THIS_REFUND_TAX_AMT()
        {
            decimal vINCOME_TAX_AMT = iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue);
            decimal vSP_TAX_AMT = iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue);

            //초기화.
            A90_THIS_REFUND_TAX_AMT.EditValue = 0;      //근로소득 당월 조정환급세액             
             
            //(9)당월 조정환금세액 
            //소득세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A90_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT;
            }
            else if (vINCOME_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A90_THIS_REFUND_TAX_AMT.EditValue = NEXT_REFUND_TAX_AMT_F();
                A90_PAY_INCOME_TAX_AMT.EditValue = vINCOME_TAX_AMT -
                                                    iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue);
            }
            else if (vINCOME_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A90_THIS_REFUND_TAX_AMT.EditValue = vINCOME_TAX_AMT;
                A90_PAY_INCOME_TAX_AMT.EditValue = 0;
            }
            //농특세
            if (NEXT_REFUND_TAX_AMT_F() <= 0)
            {
                A90_PAY_SP_TAX_AMT.EditValue = vSP_TAX_AMT;
            }
            else if (vSP_TAX_AMT > NEXT_REFUND_TAX_AMT_F())
            {
                A90_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue) +
                                                    NEXT_REFUND_TAX_AMT_F();
                A90_PAY_SP_TAX_AMT.EditValue = (vINCOME_TAX_AMT + vSP_TAX_AMT) -
                                                iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue); ;
            }
            else if (vSP_TAX_AMT < NEXT_REFUND_TAX_AMT_F())
            {
                A90_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue) +
                                                    vSP_TAX_AMT;
                A90_PAY_SP_TAX_AMT.EditValue = 0;
            }

            //합계
            TOTAL_THIS_REFUND_TAX_AMT();    //당월조정환급세액 합계 
            //TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            //TOTAL_PAY_SP_TAX_AMT();         //농특세
        }


        private void TOTAL_PERSON_CNT()
        {
            //인원수 합계 
            A99_PERSON_CNT.EditValue = 0;
            A99_PERSON_CNT.EditValue = iConv.ISDecimaltoZero(A10_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A20_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A30_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A40_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A47_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A50_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A60_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A69_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A70_PERSON_CNT.EditValue) +
                                        iConv.ISDecimaltoZero(A80_PERSON_CNT.EditValue);
        }

        private void TOTAL_PAYMENT_AMT()
        {
            //총지급액 합계 
            A99_PAYMENT_AMT.EditValue = 0;
            A99_PAYMENT_AMT.EditValue = iConv.ISDecimaltoZero(A10_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A20_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A30_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A40_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A47_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A50_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A60_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A70_PAYMENT_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A80_PAYMENT_AMT.EditValue);
        }

        private void TOTAL_INCOME_TAX_AMT()
        {
            //소득세 합계 
            A99_INCOME_TAX_AMT.EditValue = 0;
            if(iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue);
            }
            if (iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue) > 0)
            {
                A99_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A99_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue);
            } 
        }

        private void TOTAL_SP_TAX_AMT()
        {
            //농특세 합계
            A99_SP_TAX_AMT.EditValue = 0;
            A99_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A50_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A60_SP_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A90_SP_TAX_AMT.EditValue); 
        }

        private void TOTAL_ADD_TAX_AMT()
        {
            //가산세 합계
            A99_ADD_TAX_AMT.EditValue = 0;
            A99_ADD_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A20_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A40_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A47_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A50_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A60_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A69_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A70_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A80_ADD_TAX_AMT.EditValue) +
                                        iConv.ISDecimaltoZero(A90_ADD_TAX_AMT.EditValue);
        }

        private void TOTAL_THIS_REFUND_TAX_AMT()
        {
            //당월 환급세액
            //18.조정환급세액 한도내에서 적용 
            A99_THIS_REFUND_TAX_AMT.EditValue = 0;
            A99_THIS_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A20_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A40_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A47_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A50_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A60_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A69_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A70_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A80_THIS_REFUND_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A90_THIS_REFUND_TAX_AMT.EditValue);

            THIS_ADJUST_REFUND_TAX_AMT.EditValue = A99_THIS_REFUND_TAX_AMT.EditValue; 
            CAL_NEXT_REFUND_TAX_AMT();
        }

        private void TOTAL_PAY_INCOME_TAX_AMT()
        {
            //납부세액 - 소득세등(가산세 포함)
            A99_PAY_INCOME_TAX_AMT.EditValue = 0;
            A99_PAY_INCOME_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A20_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A40_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A47_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A50_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A60_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A69_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A70_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A80_PAY_INCOME_TAX_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(A90_PAY_INCOME_TAX_AMT.EditValue);
        }

        private void TOTAL_PAY_SP_TAX_AMT()
        {
            //농특세 합계 
            A99_PAY_SP_TAX_AMT.EditValue = 0;
            A99_PAY_SP_TAX_AMT.EditValue = iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue) +
                                            iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue);
        }
                
        #endregion


        // 당월 조정환급세액 및 납부세액 소득세등 자동 계산.
        private void CAL_THIS_REFUND_TAX_AMT()
        {            
            //A10. 근로소득 가감계-당월 조정 환급세액 및 납부세액 
            CAL_A10_THIS_REFUND_TAX_AMT();

            ////A20.퇴직소득-당월 조정 환급세액 및 납부세액  
            //CAL_A20_THIS_REFUND_TAX_AMT();

            ////A30.사업소득-당월 조정 환급세액 및 납부세액 
            //CAL_A30_THIS_REFUND_TAX_AMT();

            ////A40.기타소득-당월 조정 환급세액 및 납부세액 
            //CAL_A40_THIS_REFUND_TAX_AMT();
            
            ////A47.연금소득-당월 조정 환급세액 및 납부세액 
            //CAL_A47_THIS_REFUND_TAX_AMT();

            ////A50.이자소득 가감계-당월 조정 환급세액 및 납부세액 
            //CAL_A50_THIS_REFUND_TAX_AMT();

            ////A60.배당소득 가감계-당월 조정 환급세액 및 납부세액 
            //CAL_A60_THIS_REFUND_TAX_AMT();

            ////A69.저축해지 추징세액등-당월 조정 환급세액 및 납부세액 
            //CAL_A69_THIS_REFUND_TAX_AMT();

            ////A70.비거주자 양도소득-당월 조정 환급세액 및 납부세액 
            //CAL_A70_THIS_REFUND_TAX_AMT();

            ////A80.법인/낸외국법인원천-당월 조정 환급세액 및 납부세액 
            //CAL_A80_THIS_REFUND_TAX_AMT();

            //A90.수정신고세액-당월 조정 환급세액 및 납부세액 
            CAL_A90_THIS_REFUND_TAX_AMT();

            TOTAL_PAY_INCOME_TAX_AMT();     //소득세 합계
            TOTAL_PAY_SP_TAX_AMT();         //농특세
        }

        // 미환급 세액 계산.
        private void CAL_REFUND_BALANCE_AMT()
        {
            //14. 차감잔액(12-13) 
            REFUND_BALANCE_AMT.EditValue = iConv.ISDecimaltoZero(RECEIVE_REFUND_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(ALREADY_REFUND_TAX_AMT.EditValue);

            CAL_ADJUST_REFUND_TAX_AMT();
        }

        // 일반환급 
        private void CAL_GENERAL_REFUND_AMT()
        {
            //초기화.
            GENERAL_REFUND_AMT.EditValue = 0;           //일반환급 

            //A10. 근로소득 가감계-당월 조정 환급세액 및 납부세액 
            //소득세 
            if (iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A10_INCOME_TAX_AMT.EditValue));
                
            }
            //농특세   
            if (iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A10_PAY_SP_TAX_AMT.EditValue));
            }
            
            //A20.퇴직소득-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A20_INCOME_TAX_AMT.EditValue));
            }
            
            //A30.사업소득-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue));
            }
            //농특세 
            //(15)일반환급 
            if (iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue));                 
            }
           
            //A40.기타소득-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A40_INCOME_TAX_AMT.EditValue));               
            }
            
            //A47.연금소득-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A47_INCOME_TAX_AMT.EditValue));
            }
            
            //A50.이자소득 가감계-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A50_INCOME_TAX_AMT.EditValue));
            }
            //농특세  
            if (iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A50_PAY_SP_TAX_AMT.EditValue));
            }
            
            //A60.배당소득 가감계-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A60_INCOME_TAX_AMT.EditValue));
            }
            //농특세 
            if (iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A60_PAY_SP_TAX_AMT.EditValue));                 
            }
            
            //A69.저축해지 추징세액등-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A69_INCOME_TAX_AMT.EditValue));
            }
            
            //A70.비거주자 양도소득-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A70_INCOME_TAX_AMT.EditValue));
            }
            
            //A80.법인/낸외국법인원천-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A80_INCOME_TAX_AMT.EditValue));
            }
             
            //A90.수정신고세액-당월 조정 환급세액 및 납부세액 
            //납부세액-소득세등(가산세 포함) 
            if (iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A90_INCOME_TAX_AMT.EditValue));
            }
            //농특세  
            if (iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue) < 0)
            {
                GENERAL_REFUND_AMT.EditValue = Math.Abs(iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue)) +
                                                Math.Abs(iConv.ISDecimaltoZero(A90_PAY_SP_TAX_AMT.EditValue));                 
            }

            //조정대상 환급세액 반영.
            CAL_ADJUST_REFUND_TAX_AMT();
        }

        private void CAL_ADJUST_REFUND_TAX_AMT()
        {
            //18.조정대상 환급세액(14 + 15 + 16 + 17)
            ADJUST_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(REFUND_BALANCE_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(GENERAL_REFUND_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(FINANCIAL_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(ETC_REFUND_FINANCIAL_AMT.EditValue) +
                                                iConv.ISDecimaltoZero(ETC_REFUND_MERGER_AMT.EditValue);

            //19.당월조정환급세액계.
            CAL_THIS_REFUND_TAX_AMT();
            //20.차월이월환급세액.
            CAL_NEXT_REFUND_TAX_AMT();
        }
         
        private void CAL_NEXT_REFUND_TAX_AMT()
        {
            //20 차월이월 환급세액(18-19)
            NEXT_REFUND_TAX_AMT.EditValue = iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) -
                                            iConv.ISDecimaltoZero(THIS_ADJUST_REFUND_TAX_AMT.EditValue);
        }

        //차월이월 금액 계산
        private decimal NEXT_REFUND_TAX_AMT_F()
        {
            decimal mNEXT_REFUND_TAX_AMT = 0;

            TOTAL_THIS_REFUND_TAX_AMT();

            mNEXT_REFUND_TAX_AMT = iConv.ISDecimaltoZero(ADJUST_REFUND_TAX_AMT.EditValue) -
                                    iConv.ISDecimaltoZero(THIS_ADJUST_REFUND_TAX_AMT.EditValue);

            return mNEXT_REFUND_TAX_AMT;
        }

    }
}