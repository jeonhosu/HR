using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0213
{
    public partial class HRMF0213 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private string mMessageError = string.Empty;
        private string mREPORT_TYPE = string.Empty;
        private string mREPORT_FILE_NAME = string.Empty;

        private int mStartPage = 1;

        #endregion;

        #region ----- UpLoad / DownLoad Variables -----

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private ItemImageInfomationFTP mImageFTP;

        private string mFTP_Source_Directory = string.Empty;            // ftp 소스 디렉토리.
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.
        private string mClient_Target_Directory = string.Empty;         // 실제 디렉토리 
        private string mClient_ImageDirectory = string.Empty;           // 클라이언트 이미지 디렉토리.
        private string mPassive = string.Empty;                         // Passvie mode.
        private string mFileExtension = ".JPG";

        private bool mIsGetInformationFTP = false;                      // FTP 정보 상태.
        private bool mIsFormLoad = false;                               // NEWMOVE 이벤트 제어.

        #endregion;

        #region ----- initialize -----
        public HRMF0213(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.DoubleBuffered = true;
            this.Visible = false;
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mIsFormLoad = false;
        }
        #endregion

        #region ----- DATA FIND ------

        private void SEARCH_DB()
        {
            // 데이터 조회.
            idaPERSON.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idaPERSON.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);


            igrPERSON_INFO.LastConfirmChanges();
            idaPERSON.OraSelectData.AcceptChanges();
            idaPERSON.Refillable = true;

            idaPERSON.Fill();
        }

        private void isSetCommonLookUpParameter(string P_GROUP_CODE, string P_CODE_NAME, String P_USABLE_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", P_CODE_NAME);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_USABLE_YN);
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

        #endregion;



        #region ----- XLPrinting1 Methods -----

        private void XLPrinting_Main(string pOutput_Type)
        {
            string vSaveFileName = string.Empty;
            if (pOutput_Type == "FILE")
            {
                SaveFileDialog vSaveFileDialog = new SaveFileDialog();
                vSaveFileDialog.RestoreDirectory = true;
                vSaveFileDialog.Filter = "xlsx file(*.xlsx)|*.xlsx";
                vSaveFileDialog.DefaultExt = "xlsx";

                if (vSaveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    vSaveFileName = vSaveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            } 

            IDC_GET_DATE.ExecuteNonQuery();
            object vCurrDate = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_STD_DATE", vCurrDate);
            IDC_GET_REPORT_SET_P.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0213");
            IDC_GET_REPORT_SET_P.ExecuteNonQuery();
            mREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_TYPE"));
            mREPORT_FILE_NAME = iString.ISNull(IDC_GET_REPORT_SET_P.GetCommandParamValue("O_REPORT_FILE_NAME"));

            if (mREPORT_TYPE.ToUpper() == "ISV")
            {
                if (mREPORT_FILE_NAME == String.Empty)
                { 
                    mREPORT_FILE_NAME = "HRMF0213_002.xlsx";
                }
                XLPrinting_ISV(mREPORT_FILE_NAME, pOutput_Type, vSaveFileName);
            }
            else if (mREPORT_TYPE.ToUpper() == "NFV")
            {
                if (mREPORT_FILE_NAME == String.Empty)
                { 
                    mREPORT_FILE_NAME = "HRMF0213_004.xlsx";
                }
                XLPrinting_NFV(mREPORT_FILE_NAME, pOutput_Type, vSaveFileName);
            } 
            else if (mREPORT_TYPE.ToUpper() == "SIV")
            {
                if (mREPORT_FILE_NAME == string.Empty)
                {
                    mREPORT_FILE_NAME = "HRMF0213_003.xlsx";
                } 
                XLPrinting_SIV(mREPORT_FILE_NAME, pOutput_Type, vSaveFileName);
            }
            else
            {
                //-------------------------------------------------------------------------
                if (mREPORT_FILE_NAME == String.Empty)
                { 
                    mREPORT_FILE_NAME = "HRMF0213_001.xlsx";
                }
                XLPrinting(mREPORT_FILE_NAME, pOutput_Type, vSaveFileName);
            }

            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // 인쇄 완료 메시지 출력
        }

        private void XLPrinting(string pREPORT_FILE_NAME, string pOutput_Type, string pSaveFileName)
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                xlPrinting.OpenFileNameExcel = mREPORT_FILE_NAME;
                xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------

                //xlPrinting.PreView();

                // 전체 Grid
                int vTerritory1 = GetTerritory(igrPERSON_INFO.TerritoryLanguage);          // 기본사항1

                int vPageNumber = 1;
                int vRowCnt = 1;
                int vColCnt = 1;

                int vPrintRow = 1;
                int vPrintCol = 1;

                // 체크한 항목 정보
                int vIndexCheckBox = igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"); // select의 컬럼 인덱스
                int vTotalRow = igrPERSON_INFO.RowCount; // igrPERSON_INFO의 총 행수

                igrPERSON_INFO.Focus();

                //사원사진 인쇄

                System.Drawing.SizeF vSize = new System.Drawing.SizeF(86.2F, 103.9F);
                System.Drawing.PointF vPoint = new System.Drawing.PointF(20F, 50F);

                //System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                //System.Drawing.PointF vPoint = new System.Drawing.PointF(25F, 125F);
                //mPrinting.XLBarCode(pImageName, vSize, vPoint);


                if (pOutput_Type == "FILE" && pSaveFileName != string.Empty)
                {
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(pSaveFileName);
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
                    vMessageText = string.Format(" Writing Starting...");
                }
                else
                {
                    vMessageText = string.Format(" Printing Starting...");
                }

                int vIDX_PERSON_NUM = igrPERSON_INFO.GetColumnToIndex("PERSON_NUM");
                for (int nRow = 0; nRow < vTotalRow; nRow++)
                {
                    if ((string)igrPERSON_INFO.GetCellValue(nRow, vIndexCheckBox) == "Y") // 선택 항목에 체크되어진 것만 출력하기 위한 조건
                    {
                        // Main이 되는 igrPERSON_INFO Grid의 Line 단위 정보 확인
                        igrPERSON_INFO.CurrentCellMoveTo(nRow, vIndexCheckBox);

                        // 증명사진 Image 경로 및 파일명
                        string sDownLoadFile = string.Empty;
                        try
                        {
                            string vPersonNumber = iString.ISNull(igrPERSON_INFO.GetCellValue(nRow, vIDX_PERSON_NUM));
                            sDownLoadFile = isViewItemImage(vPersonNumber);
                        }
                        catch
                        {

                        }

                        if (3 < vColCnt)  //한 row에 3건 인쇄위해.
                        {
                            vPrintCol = 1;
                            vPrintRow = vPrintRow + 20;
                            vColCnt = 1;
                        }

                        // 9건 이상일 경우, 페이지 스킵 
                        if (vRowCnt > 9)
                        {
                            vPageNumber = vPageNumber + 1;
                            vRowCnt = 1;
                        }

                        //출력부분 
                        xlPrinting.XLWirte(igrPERSON_INFO, nRow, vTerritory1, sDownLoadFile, vPrintRow, vPrintCol, vSize, vPoint);

                        vPrintCol = vPrintCol + 15;
                        vRowCnt = vRowCnt + 1;
                        vColCnt = vColCnt + 1;
                    }
                }
                if (pOutput_Type == "FILE")
                {
                    xlPrinting.Save(pSaveFileName);
                }
                else if (pOutput_Type == "PRINT")
                {
                    xlPrinting.Printing(mStartPage, vPageNumber);
                }

                //xlPrinting.PreView();
                xlPrinting.Dispose();
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        private void XLPrinting_ISV(string pREPORT_FILE_NAME, string pOutput_Type, string pSaveFileName)
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = mREPORT_FILE_NAME;
                xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------

                //xlPrinting.PreView();

                // 전체 Grid
                int vTerritory1 = GetTerritory(igrPERSON_INFO.TerritoryLanguage);          // 기본사항1

                int vPageNumber = 1;
                int vRowCnt = 1;
                int vColCnt = 1;

                int vPrintRow = 1;
                int vPrintCol = 1;

                // 체크한 항목 정보
                int vIndexCheckBox = igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"); // select의 컬럼 인덱스
                int vTotalRow = igrPERSON_INFO.RowCount; // igrPERSON_INFO의 총 행수

                igrPERSON_INFO.Focus();

                //사원사진 인쇄

                System.Drawing.SizeF vSize = new System.Drawing.SizeF(86.2F, 103.9F);
                System.Drawing.PointF vPoint = new System.Drawing.PointF(20F, 50F);

                //System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                //System.Drawing.PointF vPoint = new System.Drawing.PointF(25F, 125F);
                //mPrinting.XLBarCode(pImageName, vSize, vPoint);


                if (pOutput_Type == "FILE" && pSaveFileName != string.Empty)
                {
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(pSaveFileName);
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
                    vMessageText = string.Format(" Writing Starting...");
                }
                else
                {
                    vMessageText = string.Format(" Printing Starting...");
                }

                int vIDX_PERSON_NUM = igrPERSON_INFO.GetColumnToIndex("PERSON_NUM");
                for (int nRow = 0; nRow < vTotalRow; nRow++)
                {
                    if ((string)igrPERSON_INFO.GetCellValue(nRow, vIndexCheckBox) == "Y") // 선택 항목에 체크되어진 것만 출력하기 위한 조건
                    {
                        // Main이 되는 igrPERSON_INFO Grid의 Line 단위 정보 확인
                        igrPERSON_INFO.CurrentCellMoveTo(nRow, vIndexCheckBox);

                        // 증명사진 Image 경로 및 파일명
                        string sDownLoadFile = string.Empty;
                        try
                        {
                            string vPersonNumber = iString.ISNull(igrPERSON_INFO.GetCellValue(nRow, vIDX_PERSON_NUM));
                            sDownLoadFile = isViewItemImage(vPersonNumber);
                        }
                        catch
                        {

                        }

                        if (3 < vColCnt)  //한 row에 3건 인쇄위해.
                        {
                            vPrintCol = 1;
                            vPrintRow = vPrintRow + 20;
                            vColCnt = 1;
                        }

                        // 9건 이상일 경우, 페이지 스킵 
                        if (vRowCnt > 9)
                        {
                            vPageNumber = vPageNumber + 1;
                            vRowCnt = 1;
                        }

                        //출력부분 
                        xlPrinting.XLWirte(igrPERSON_INFO, nRow, vTerritory1, sDownLoadFile, vPrintRow, vPrintCol, vSize, vPoint);

                        vPrintCol = vPrintCol + 15;
                        vRowCnt = vRowCnt + 1;
                        vColCnt = vColCnt + 1;
                    }
                }
                if (pOutput_Type == "FILE")
                {
                    xlPrinting.Save(pSaveFileName);
                }
                else if (pOutput_Type == "PRINT")
                {
                    xlPrinting.Printing(mStartPage, vPageNumber);
                }

                //xlPrinting.PreView();
                xlPrinting.Dispose();
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        private void XLPrinting_NFV(string pREPORT_FILE_NAME, string pOutput_Type, string pSaveFileName)
        {
            string vMessageText = string.Empty;

            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = pREPORT_FILE_NAME; 
                xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------

                //xlPrinting.PreView();

                // 전체 Grid
                int vTerritory1 = GetTerritory(igrPERSON_INFO.TerritoryLanguage);          // 기본사항1
                int vPageNumber = 1;
                int vRowCnt = 1;
                int vColCnt = 1;

                int vPrintRow = 1;
                int vPrintCol = 1;

                // 체크한 항목 정보
                int vIndexCheckBox = igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"); // select의 컬럼 인덱스
                int vTotalRow = igrPERSON_INFO.RowCount; // igrPERSON_INFO의 총 행수

                igrPERSON_INFO.Focus();

                //사원사진 인쇄

                System.Drawing.SizeF vSize = new System.Drawing.SizeF(86.2F, 103.9F);
                System.Drawing.PointF vPoint = new System.Drawing.PointF(20F, 50F);

                //System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                //System.Drawing.PointF vPoint = new System.Drawing.PointF(25F, 125F);
                //mPrinting.XLBarCode(pImageName, vSize, vPoint);


                if (pOutput_Type == "FILE" && pSaveFileName != string.Empty)
                {
                   System.IO.FileInfo vFileName = new System.IO.FileInfo(pSaveFileName);
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
                    vMessageText = string.Format(" Writing Starting...");
                }
                else
                {
                    vMessageText = string.Format(" Printing Starting...");
                }

                int vIDX_PERSON_NUM = igrPERSON_INFO.GetColumnToIndex("PERSON_NUM");
                for (int nRow = 0; nRow < vTotalRow; nRow++)
                {
                    if ((string)igrPERSON_INFO.GetCellValue(nRow, vIndexCheckBox) == "Y") // 선택 항목에 체크되어진 것만 출력하기 위한 조건
                    {
                        // Main이 되는 igrPERSON_INFO Grid의 Line 단위 정보 확인
                        igrPERSON_INFO.CurrentCellMoveTo(nRow, vIndexCheckBox);

                        // 증명사진 Image 경로 및 파일명
                        string sDownLoadFile = string.Empty;
                        try
                        {
                            string vPersonNumber = iString.ISNull(igrPERSON_INFO.GetCellValue(nRow, vIDX_PERSON_NUM));
                            sDownLoadFile = isViewItemImage(vPersonNumber);
                        }
                        catch
                        {

                        }

                        if (3 < vColCnt)  //한 row에 3건 인쇄위해.
                        {
                            vPrintCol = 1;
                            vPrintRow = vPrintRow + 20;
                            vColCnt = 1;
                        }

                        // 9건 이상일 경우, 페이지 스킵 
                        if (vRowCnt > 9)
                        {
                            vPageNumber = vPageNumber + 1;
                            vRowCnt = 1;
                        }

                        //출력부분 
                        xlPrinting.XLWirte(igrPERSON_INFO, nRow, vTerritory1, sDownLoadFile, vPrintRow, vPrintCol, vSize, vPoint);

                        vPrintCol = vPrintCol + 15;
                        vRowCnt = vRowCnt + 1;
                        vColCnt = vColCnt + 1;
                    }
                }
                if (pOutput_Type == "FILE")
                {
                    xlPrinting.Save(pSaveFileName);
                }
                else if (pOutput_Type == "PRINT")
                {
                    xlPrinting.Printing(mStartPage, vPageNumber);
                }

                //xlPrinting.PreView();
                xlPrinting.Dispose();
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }
        
        private void XLPrinting_SIV(string pREPORT_FILE_NAME, string pOutput_Type, string pSaveFileName)
        {
            string vMessageText = string.Empty;
            XLPrinting xlPrinting = new XLPrinting();

            try
            {
                //-------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = pREPORT_FILE_NAME;
                xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------

                //xlPrinting.PreView();

                // 전체 Grid
                int vTerritory1 = GetTerritory(igrPERSON_INFO.TerritoryLanguage);          // 기본사항1
                int vPageNumber = 1;
                int vRowCnt = 1;
                int vColCnt = 1;

                int vPrintRow = 1;
                int vPrintCol = 1;

                // 체크한 항목 정보
                int vIndexCheckBox = igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"); // select의 컬럼 인덱스
                int vTotalRow = igrPERSON_INFO.RowCount; // igrPERSON_INFO의 총 행수

                igrPERSON_INFO.Focus();

                //사원사진 인쇄

                System.Drawing.SizeF vSize = new System.Drawing.SizeF(86.2F, 103.9F);
                System.Drawing.PointF vPoint = new System.Drawing.PointF(20F, 50F);

                //System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                //System.Drawing.PointF vPoint = new System.Drawing.PointF(25F, 125F);
                //mPrinting.XLBarCode(pImageName, vSize, vPoint);


                if (pOutput_Type == "FILE" && pSaveFileName != string.Empty)
                {
                    System.IO.FileInfo vFileName = new System.IO.FileInfo(pSaveFileName);
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
                    vMessageText = string.Format(" Writing Starting...");
                }
                else
                {
                    vMessageText = string.Format(" Printing Starting...");
                }

                int vIDX_PERSON_NUM = igrPERSON_INFO.GetColumnToIndex("PERSON_NUM");
                for (int nRow = 0; nRow < vTotalRow; nRow++)
                {
                    if ((string)igrPERSON_INFO.GetCellValue(nRow, vIndexCheckBox) == "Y") // 선택 항목에 체크되어진 것만 출력하기 위한 조건
                    {
                        // Main이 되는 igrPERSON_INFO Grid의 Line 단위 정보 확인
                        igrPERSON_INFO.CurrentCellMoveTo(nRow, vIndexCheckBox);

                        // 증명사진 Image 경로 및 파일명
                        string sDownLoadFile = string.Empty;
                        try
                        {
                            string vPersonNumber = iString.ISNull(igrPERSON_INFO.GetCellValue(nRow, vIDX_PERSON_NUM));
                            sDownLoadFile = isViewItemImage(vPersonNumber);
                        }
                        catch
                        {

                        }

                        if (3 < vColCnt)  //한 row에 3건 인쇄위해.
                        {
                            vPrintCol = 1;
                            vPrintRow = vPrintRow + 20;
                            vColCnt = 1;
                        }

                        // 9건 이상일 경우, 페이지 스킵 
                        if (vRowCnt > 9)
                        {
                            vPageNumber = vPageNumber + 1;
                            vRowCnt = 1;
                        }

                        //출력부분 
                        xlPrinting.XLWirte(igrPERSON_INFO, nRow, vTerritory1, sDownLoadFile, vPrintRow, vPrintCol, vSize, vPoint);

                        vPrintCol = vPrintCol + 15;
                        vRowCnt = vRowCnt + 1;
                        vColCnt = vColCnt + 1;
                    }
                }
                if (pOutput_Type == "FILE")
                {
                    xlPrinting.Save(pSaveFileName);
                }
                else if (pOutput_Type == "PRINT")
                {
                    xlPrinting.Printing(mStartPage, vPageNumber);
                }

                //xlPrinting.PreView();
                xlPrinting.Dispose();
            }
            catch (System.Exception ex)
            {
                string vMessage = ex.Message;
                xlPrinting.Dispose();
            }
        }

        #endregion;
         

        #region ----- XLPrinting1 Methods -----



        #endregion;

        #region --- Application_MainButtonClick ---

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
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
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    XLPrinting_Main("PRINT"); // 출력 함수 호출 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    XLPrinting_Main("FILE");
                }

            }
        }

        #endregion

        #region ----- Form Event -----

        private void HRMF0213_Load(object sender, EventArgs e)
        {
            this.Visible = true;
            mIsFormLoad = true;

            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();

            W_CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME_0.BringToFront();

            // 그리드 부분 업데이트 처리 위함.
            idaPERSON.FillSchema();
        }


        private void HRMF0213_Shown(object sender, EventArgs e)
        {
            STAT_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = DateTime.Today;


            mIsGetInformationFTP = GetInfomationFTP();
            if (mIsGetInformationFTP == true)
            {
                MakeDirectory();
                FTPInitializtion();
            }
            mIsFormLoad = false;
        }

        #endregion

        #region ----- Adapter Event -----

        private bool isDelete_Validate(string pTabPage)
        {
            bool ibReturn_Value = false;
            if (pTabPage == "itpPERSON_MASTER")
            {
                ibReturn_Value = false;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);   // 사원정보 삭제 불가.
            }
            return ibReturn_Value;
        }




        #endregion

        #region ----- idaPERSON NewRowMoved Event -----

        private void idaPERSON_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (mIsFormLoad == true)
            {
                return;
            }

            isViewItemImage(iString.ISNull(pBindingManager.DataRow["PERSON_NUM"]));
        }
        #endregion

        #region ----- lookup adapter event -----

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EMPLOYE_TYPE", null, "Y");
        }

        private void ilaEMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EMPLOYE_TYPE", null, "Y");
        }

        #endregion

        #region ----- is View Item Image Method ----

        private string isViewItemImage(string pPERSON_NUM)
        {
            if (mIsFormLoad == true)
            {
                return null;
            }

            //bool isView = false;
            string vDownLoadFile = string.Empty;
            string vTargetFileName = string.Format("{0}{1}", pPERSON_NUM.ToUpper(), mFileExtension);

            bool isDown = DownLoadItem(vTargetFileName);
            if (isDown == true)
            {
                vDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, vTargetFileName);
            }
            return vDownLoadFile;
        }

        #endregion;

        #region ----- Make Directory ----

        private void MakeDirectory()
        {
            System.IO.DirectoryInfo vClient_ImageDirectory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
            if (vClient_ImageDirectory.Exists == false) //있으면 True, 없으면 False
            {
                vClient_ImageDirectory.Create();
            }
        }

        #endregion;

        #region ----- Image View ----

        private bool ImageView(string pFileName)
        {
            bool isView = false;

            return isView;
        }

        #endregion;


        #region ----- Get Information FTP Methods -----

        private bool GetInfomationFTP()
        {
            bool isGet = false;
            try
            {
                IDC_FTP_INFO.SetCommandParamValue("W_FTP_CODE", "PERSON_PIC");
                IDC_FTP_INFO.ExecuteNonQuery();
                mImageFTP = new ItemImageInfomationFTP();

                mImageFTP.Host = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mImageFTP.Port = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mImageFTP.UserID = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_NO"));
                mImageFTP.Password = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mImageFTP.Passive_Flag = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));

                mFTP_Source_Directory = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mClient_Target_Directory = iString.ISNull(IDC_FTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));

                mClient_ImageDirectory = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Target_Directory);

                if (mImageFTP.Host != string.Empty)
                {
                    isGet = true;
                } 
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
            return isGet;
        }

        #endregion;

        #region ----- FTP Initialize -----

        private void FTPInitializtion()
        {
            mFileTransferAdv = new ISFileTransferAdv();
            mFileTransferAdv.Host = mImageFTP.Host;
            mFileTransferAdv.Port = mImageFTP.Port;
            mFileTransferAdv.UserId = mImageFTP.UserID;
            mFileTransferAdv.Password = mImageFTP.Password;
            if (mImageFTP.Passive_Flag == "Y")
            {
                mFileTransferAdv.UsePassive = true;
            }
            else
            {
                mFileTransferAdv.UsePassive = false;
            }
        }

        #endregion;

        #region ----- Image Upload Methods -----

        private bool UpLoadItem()
        {
            bool isUp = false;

            openFileDialog1.FileName = string.Format("*{0}", mFileExtension);
            openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", mFileExtension, mFileExtension);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string vChoiceFileFullPath = openFileDialog1.FileName;
                    string vChoiceFilePath = vChoiceFileFullPath.Substring(0, vChoiceFileFullPath.LastIndexOf(@"\"));
                    string vChoiceFileName = vChoiceFileFullPath.Substring(vChoiceFileFullPath.LastIndexOf(@"\") + 1);

                    mFileTransferAdv.ShowProgress = true;
                    //--------------------------------------------------------------------------------

                    string vSourceFileName = vChoiceFileName;

                    string vTargetFileName = igrPERSON_INFO.GetCellValue("PERSON_NUM") as string;
                    vTargetFileName = string.Format("{0}{1}", vTargetFileName.ToUpper(), mFileExtension);

                    mFileTransferAdv.SourceDirectory = vChoiceFilePath;
                    mFileTransferAdv.SourceFileName = vSourceFileName;
                    mFileTransferAdv.TargetDirectory = mFTP_Source_Directory;
                    mFileTransferAdv.TargetFileName = vTargetFileName;

                    bool isUpLoad = mFileTransferAdv.Upload();

                    if (isUpLoad == true)
                    {
                        isUp = true;
                        bool isView = ImageView(vChoiceFileFullPath);
                    }
                    else
                    {
                    }
                }
                catch
                {
                }
            }
            System.IO.Directory.SetCurrentDirectory(mClient_Base_Path);
            return isUp;
        }


        private void ibtPERSON_PICTURE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (mIsGetInformationFTP == true)
            {
                UpLoadItem();
            }
        }


        #endregion;

        #region ----- Image Download Methods -----

        private bool DownLoadItem(string pFileName)
        {
            bool isDown = false;

            string vSourceDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, pFileName);
            string vTargetDownLoadFile = string.Format("{0}\\_{1}", mClient_ImageDirectory, pFileName);

            string vBeforeSourceFileName = string.Format("{0}", pFileName);
            string vBeforeTargetFileName = string.Format("_{0}", pFileName);

            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransferAdv.SourceDirectory = mFTP_Source_Directory;
            mFileTransferAdv.SourceFileName = vBeforeSourceFileName;
            mFileTransferAdv.TargetDirectory = mClient_ImageDirectory;
            mFileTransferAdv.TargetFileName = vBeforeTargetFileName;

            isDown = mFileTransferAdv.Download();

            if (isDown == true)
            {
                try
                {
                    System.IO.File.Delete(vSourceDownLoadFile);
                    System.IO.File.Move(vTargetDownLoadFile, vSourceDownLoadFile);

                    isDown = true;
                }
                catch
                {
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTargetDownLoadFile);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore
                    }
                }
            }
            else
            {
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTargetDownLoadFile);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore
                }
            }

            return isDown;
        }

        #endregion;

        private void iedNAME_0_KeyUp(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        // 전체선택 버튼
        private void btnSELECT_ALL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrPERSON_INFO.RowCount; i++)
            {
                igrPERSON_INFO.SetCellValue(i, igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"), "Y");
            }
        }

        // 취소 버튼
        private void btnCONFIRM_CANCEL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrPERSON_INFO.RowCount; i++)
            {
                igrPERSON_INFO.SetCellValue(i, igrPERSON_INFO.GetColumnToIndex("SELECT_CHECK_YN"), "N");
            }
        }

        private void HRMF0213_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (mIsGetInformationFTP == true)
            {
                System.IO.DirectoryInfo vClient_ImageDirectory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
                if (vClient_ImageDirectory.Exists == true)
                {
                    try
                    {
                        vClient_ImageDirectory.Delete(true);
                    }
                    catch
                    {
                    }
                }
            }
        }

        #region ----- Lookup Event -----

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnabled_YN);
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ilaOPERATING_UNIT_0_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

    }

    #region ----- User Make Class -----

    public class ItemImageInfomationFTP
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = string.Empty;
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public ItemImageInfomationFTP()
        {
        }

        public ItemImageInfomationFTP(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mPassive_Flag = pPassive_Flag;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string Passive_Flag
        {
            get
            {
                return mPassive_Flag;
            }
            set
            {
                mPassive_Flag = value;
            }
        }
        #endregion;
    }

    #endregion;

}