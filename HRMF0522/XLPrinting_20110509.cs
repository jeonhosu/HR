using System;
using ISCommonUtil;

namespace HRMF0522
{
    /// <summary>
    /// XLPrint Class를 이용해 Report물 제어 
    /// </summary>
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGridAdvEx;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private string mXLOpenFileName = string.Empty;

        private int[] mIndexGridColumns = new int[0] { };

        private int mPositionPrintLineSTART = 1; //내용 출력시 엑셀 시작 행 위치 지정
        private int[] mIndexXLWriteColumn = new int[0] { }; //엑셀에 출력할 열 위치 지정

        private int mMaxIncrement = 45; //실제 출력되는 행의 시작부터 끝 행의 범위
        private int mSumPrintingLineCopy = 1; //엑셀의 선택된 쉬트에 복사되어질 시작 행 위치 및 누적 행 값
        private int mMaxIncrementCopy = 70; //반복 복사되어질 행의 최대 범위

        private int mXLColumnAreaSTART = 1; //복사되어질 쉬트의 폭, 시작열
        private int mXLColumnAreaEND = 45;  //복사되어질 쉬트의 폭, 종료열

        #endregion;

        #region ----- Property -----

        /// <summary>
        /// 모든 Error Message 출력
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        /// <summary>
        /// Message 출력할 Grid
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISGridAdvEx MessageGridEx
        {
            set
            {
                mGridAdvEx = value;
            }
        }

        /// <summary>
        /// 전체 Data 진행 ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar1
        {
            set
            {
                mProgressBar1 = value;
            }
        }

        /// <summary>
        /// Page당 Data 진행 ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar2
        {
            set
            {
                mProgressBar2 = value;
            }
        }

        /// <summary>
        /// Ope할 Excel File 이름
        /// </summary>
        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting()
        {
            mPrinting = new XL.XLPrint();
        }

        #endregion;

        #region ----- Interior Use Methods ----

        #region ----- MessageGrid Methods ----

        private void MessageGrid(string pMessage)
        {
            int vCountRow = mGridAdvEx.RowCount;
            vCountRow = vCountRow + 1;
            mGridAdvEx.RowCount = vCountRow;

            int vCurrentRow = vCountRow - 1;

            mGridAdvEx.SetCellValue(vCurrentRow, 0, pMessage);

            mGridAdvEx.CurrentCellMoveTo(vCurrentRow, 0);
            mGridAdvEx.Focus();
            mGridAdvEx.CurrentCellActivate(vCurrentRow, 0);
        }

        #endregion;

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 3;
                int vSkip = vFileName.Length - vCutRight;
                string vTextNumber = vFileName.Substring(vSkip, vCutRight);
                int vNumber = int.Parse(vTextNumber);

                if (vNumber > vMaxNumber)
                {
                    vMaxNumber = vNumber;
                }
            }

            return vMaxNumber;
        }

        #endregion;

        #endregion;

        #region ----- XLPrint Define Methods ----

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
        }

        #endregion;

        #region ----- XL File Open -----

        public bool XLFileOpen()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mPrinting.XLOpenFile(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlAllLineClear(XL.XLPrint pPrinting)
        {
            int vXLColumn1 = 2;  //No[OPERATION_SEQ_NO]
            int vXLColumn2 = 4;  //공정명[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //공정 진행시 작업 조건[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //선그리기, 시작 열
            int vXLDrawLineColumnEND = 45;  //선그리기, 종료 열

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            //pPrinting.XLActiveSheet(2);
            pPrinting.XLActiveSheet("SourceTab1");

            for (int vXLine = mPositionPrintLineSTART; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //공정명[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //공정 진행시 작업 조건[OPERATION_COMMENT]

                if (vXLine < mMaxIncrementCopy)
                {
                    pPrinting.XL_LineClear(vXLine, vXLDrawLineColumnSTART, vXLDrawLineColumnEND);
                }
            }
        }

        #endregion;

        #region ----- Line Clear Methods ----

        //XlLineClear(mPrinting, vPrintingLine);
        private void XlLineClear(XL.XLPrint pPrinting, int pPrintingLine)
        {
            int vXLColumn1 = 2;  //No[OPERATION_SEQ_NO]
            int vXLColumn2 = 4;  //공정명[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //공정 진행시 작업 조건[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //선그리기, 시작 열
            int vXLDrawLineColumnEND = 45;  //선그리기, 종료 열

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            for (int vXLine = pPrintingLine; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //공정명[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //공정 진행시 작업 조건[OPERATION_COMMENT]

                if (vXLine < mMaxIncrementCopy)
                {
                    pPrinting.XL_LineClear(vXLine, vXLDrawLineColumnSTART, vXLDrawLineColumnEND);
                }
            }
        }

        #endregion;

        #region ----- Define Print Column Methods ----

        private void XLDefinePrintColumn(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            try
            {
                //Grid의 [Edit] 상의 [DataColumn] 열에 있는 열 이름을 지정 하면 된다.
                string[] vGridDataColumns = new string[]
                {
                    "NAME",
                    "PERSON_NUM",
                    "DEPT_NAME",
                    "POST_NAME",
                    "JOB_CLASS_NAME",
                    "SUPPLY_DATE",
                    "BANK_NAME",
                    "BANK_ACCOUNTS",
                    "REAL_AMOUNT"
                };

                int vIndexColumn = 0;
                mIndexGridColumns = new int[vGridDataColumns.Length];

                foreach (string vName in vGridDataColumns)
                {
                    mIndexGridColumns[vIndexColumn] = pGrid.GetColumnToIndex(vName);
                    vIndexColumn++;
                }

                //엑셀에 출력될 열 위치 지정
                int[] vXLColumns = new int[]
                {
                    28,
                    28,
                    28,
                    29,
                    29,
                    29,
                    30,
                    30,
                    60
                };
                mIndexXLWriteColumn = new int[vXLColumns.Length];
                for (int vCol = 0; vCol < vXLColumns.Length; vCol++)
                {
                    mIndexXLWriteColumn[vCol] = vXLColumns[vCol];
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print HeaderColumns Methods ----

        private void XLHeaderColumns(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory, int pXLine)
        {
            int vXLine = pXLine - 1; //mPositionPrintLineSTART - 1, 출력될 내용의 행 위치에서 한행 위에 있으므로 1을 뺀다.
            int vCountColumn = mIndexGridColumns.Length;

            object vObject = null;
            int vGetIndexGridColumn = 0;

            try
            {
                if (mIndexGridColumns.Length < 1)
                {
                    return;
                }

                //Header Columns
                for (int vCol = 0; vCol < vCountColumn; vCol++)
                {
                    vGetIndexGridColumn = mIndexGridColumns[vCol];
                    switch (pTerritory)
                    {
                        case 1: //Default
                            vObject = pGrid.GridAdvExColElement[vGetIndexGridColumn].HeaderElement[0].Default;
                            mPrinting.XLSetCell(vXLine, mIndexXLWriteColumn[vCol], vObject);
                            break;
                        case 2: //KR
                            vObject = pGrid.GridAdvExColElement[vGetIndexGridColumn].HeaderElement[0].TL1_KR;
                            mPrinting.XLSetCell(vXLine, mIndexXLWriteColumn[vCol], vObject);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print Content Write Methods ----

        private object ConvertDateTime(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        //string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeLong = vDateTime.ToString("yyyy년 MM월 dd일", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        #endregion

        #region ----- New Page iF Methods ----

        private int NewPage(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTotalRow, int pSumWriteLine)
        {
            int vPrintingRowSTART = 0;
            int vPrintingRowEND = 0;

            try
            {
                vPrintingRowSTART = pSumWriteLine;
                pSumWriteLine = pSumWriteLine + mMaxIncrement;
                vPrintingRowEND = pSumWriteLine;

                //XLContentWrite(mPrinting, pGrid, pTotalRow, mPositionPrintLineSTART, mIndexXLWriteColumn, vPrintingRowSTART, vPrintingRowEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return pSumWriteLine;
        }

        #endregion;

        #region ----- Excel Clear -----

        private void XLContentClear()
        {
            mPrinting.XLActiveSheet("SourceTab1");
    
            // 첫장 지급구분/부서/직위/사번/이름
            mPrinting.XLSetCell(11, 8, "");   //부서
            mPrinting.XLSetCell(13, 8, "");   //직위
            mPrinting.XLSetCell(15, 8, "");   //사번
            mPrinting.XLSetCell(17, 8, "");   //이름                    

            // 인적사항.
            mPrinting.XLSetCell(26, 9, "");   //성명
            mPrinting.XLSetCell(26, 22, "");  //사번
            mPrinting.XLSetCell(26, 36, "");  //부서
            mPrinting.XLSetCell(27, 9, "");   //직위
            mPrinting.XLSetCell(27, 22, "");  //직군                    
            mPrinting.XLSetCell(27, 36, "");  //지급일
            mPrinting.XLSetCell(28, 9, "");   //입금은행
            mPrinting.XLSetCell(28, 22, "");  //입금계좌

            //============================================================================================
            // 기본급
            //============================================================================================
            mPrinting.XLSetCell(38, 32, "");
            mPrinting.XLSetCell(38, 36, "");
            mPrinting.XLSetCell(38, 40, "");
            mPrinting.XLSetCell(55, 15, "");
            mPrinting.XLSetCell(54, 34, "");
            mPrinting.XLSetCell(55, 34, "");
            mPrinting.XLSetCell(62, 15, "");
            mPrinting.XLSetCell(61, 34, "");
            mPrinting.XLSetCell(62, 34, "");
            mPrinting.XLSetCell(63, 15, "");
            mPrinting.XLSetCell(63, 34, "");
            mPrinting.XLSetCell(64, 25, "");
            mPrinting.XLSetCell(1, 4, "");
            mPrinting.XLSetCell(67, 4, "");  //비고
            
            // 월급여 지급항목.
            for (int nRow = 0; nRow <= 12; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 6, "");
                mPrinting.XLSetCell(42 + nRow, 15, "");
            }
            // 월급여 지급항목.
            for (int nRow = 0; nRow <= 11; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 25, "");
                mPrinting.XLSetCell(42 + nRow, 34, "");
            }
            
            //============================================================================================
            // 연장(평일)
            //============================================================================================
            mPrinting.XLSetCell(32, 12, "");
            mPrinting.XLSetCell(32, 16, "");
            mPrinting.XLSetCell(32, 20, "");
            mPrinting.XLSetCell(33, 8, "");
            mPrinting.XLSetCell(33, 12, "");
            mPrinting.XLSetCell(33, 16, "");
            mPrinting.XLSetCell(34, 8, "");
            mPrinting.XLSetCell(34, 12, "");
            mPrinting.XLSetCell(34, 16, "");
            mPrinting.XLSetCell(38, 4, "");
            mPrinting.XLSetCell(38, 8, "");
            mPrinting.XLSetCell(38, 12, "");
            mPrinting.XLSetCell(38, 16, "");
            mPrinting.XLSetCell(38, 20, "");
            mPrinting.XLSetCell(38, 24, "");
            mPrinting.XLSetCell(38, 28, "");

            // 상여 지급항목.
            for (int nRow = 0; nRow <= 5; nRow++)
            {
                mPrinting.XLSetCell(56 + nRow, 6, "");
                mPrinting.XLSetCell(56 + nRow, 15, "");
            }
            // 상여 지급항목.
            for (int nRow = 0; nRow <= 4; nRow++)
            {
                mPrinting.XLSetCell(56 + nRow, 25, "");
                mPrinting.XLSetCell(56 + nRow, 34, "");
            }
        }

        private void XLContentClear2()
        {
            mPrinting.XLActiveSheet("SourceTab2");

            // 첫장 지급구분/부서/직위/사번/이름
            mPrinting.XLSetCell(11, 8, "");   //부서
            mPrinting.XLSetCell(13, 8, "");   //직위
            mPrinting.XLSetCell(15, 8, "");   //사번
            mPrinting.XLSetCell(17, 8, "");   //이름                    

            // 인적사항.
            mPrinting.XLSetCell(26, 9, "");   //성명
            mPrinting.XLSetCell(26, 22, "");  //사번
            mPrinting.XLSetCell(26, 36, "");  //부서
            mPrinting.XLSetCell(27, 9, "");   //직위
            mPrinting.XLSetCell(27, 22, "");  //직군                    
            mPrinting.XLSetCell(27, 36, "");  //지급일
            mPrinting.XLSetCell(28, 9, "");   //입금은행
            mPrinting.XLSetCell(28, 22, "");  //입금계좌

            //============================================================================================
            // 기본급
            //============================================================================================
            mPrinting.XLSetCell(38, 32, "");
            mPrinting.XLSetCell(38, 36, "");
            mPrinting.XLSetCell(38, 40, "");
            mPrinting.XLSetCell(61, 15, "");
            mPrinting.XLSetCell(61, 34, "");
            mPrinting.XLSetCell(62, 25, "");            
            mPrinting.XLSetCell(1, 4, "");
            mPrinting.XLSetCell(65, 4, "");  //비고

            // 월급여 지급/공제항목.
            for (int nRow = 0; nRow <= 18; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 6, "");
                mPrinting.XLSetCell(42 + nRow, 15, "");
                mPrinting.XLSetCell(42 + nRow, 25, "");
                mPrinting.XLSetCell(42 + nRow, 34, "");
            }

            //============================================================================================
            // 연장(평일)
            //============================================================================================
            mPrinting.XLSetCell(32, 12, "");
            mPrinting.XLSetCell(32, 16, "");
            mPrinting.XLSetCell(32, 20, "");
            mPrinting.XLSetCell(33, 8, "");
            mPrinting.XLSetCell(33, 12, "");
            mPrinting.XLSetCell(33, 16, "");
            mPrinting.XLSetCell(34, 8, "");
            mPrinting.XLSetCell(34, 12, "");
            mPrinting.XLSetCell(34, 16, "");
            mPrinting.XLSetCell(38, 4, "");
            mPrinting.XLSetCell(38, 8, "");
            mPrinting.XLSetCell(38, 12, "");
            mPrinting.XLSetCell(38, 16, "");
            mPrinting.XLSetCell(38, 20, "");
            mPrinting.XLSetCell(38, 24, "");
            mPrinting.XLSetCell(38, 28, "");
        }

        #endregion

        #region ----- XLContent Write -----
		 
        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, int pAllowance_Row, int nAllowance_Column)
        {
            decimal vAMOUNT = 0;
            decimal vDUTY_TIME = 0;
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");
                if (pCnt == 1)
                {                    
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("NAME");                   //성명
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("PERSON_NUM");             //사번
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DEPT_NAME");              //부서
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("POST_NAME");              //직위
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("JOB_CLASS_NAME");         //직군
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("SUPPLY_DATE");            //지급일
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("BANK_NAME");              //입금은행
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");          //입금계좌                 
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("BASIC_AMOUNT");          //기본급
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("BASIC_TIME_AMOUNT");     //시급
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("DESCRIPTION");           //비고
                    // '비고'는 후에 추가로 삽입된 것이라 Column 순서가 15로 된 것임.

                    // 명세서 Report 상단에 출력될 내용
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT"); //통상시급
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("WAGE_TYPE");             //급상여구분명
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("PAY_YYYYMM");            //지급년월                    

                    int vIndexWageType = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //지급구분
                    
                    int vIDX_TOT_REAL = pGrid.GetColumnToIndex("REAL_AMOUNT");                // 총 실지급액
                    int vIDX_TOT_SUPP = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          // 총지급액
                    int vIDX_TOT_DED = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");              // 총공제액

                    int vIDX_PAY_REAL = pGrid.GetColumnToIndex("REAL_PAY_AMOUNT");            // 급여 실지급액
                    int vIDX_PAY_SUPP = pGrid.GetColumnToIndex("TOT_PAY_SUP_AMOUNT");         // 급여 총지급액
                    int vIDX_PAY_DED = pGrid.GetColumnToIndex("TOT_PAY_DED_AMOUNT");          // 급여 총공제액

                    int vIDX_BONUS_REAL = pGrid.GetColumnToIndex("REAL_BONUS_AMOUNT");        // 급여 실지급액
                    int vIDX_BONUS_SUPP = pGrid.GetColumnToIndex("TOT_BONUS_SUP_AMOUNT");     // 급여 총지급액
                    int vIDX_BONUS_DED = pGrid.GetColumnToIndex("TOT_BONUS_DED_AMOUNT");      // 급여 총공제액


                    // 첫장 지급구분/부서/직위/사번/이름
                    mPrinting.XLSetCell(11, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));   //부서
                    mPrinting.XLSetCell(13, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //직위
                    mPrinting.XLSetCell(15, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));   //사번
                    mPrinting.XLSetCell(17, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //이름                    

                    // 인적사항.
                    mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //성명
                    mPrinting.XLSetCell(26, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));  //사번
                    mPrinting.XLSetCell(26, 36, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));  //부서
                    mPrinting.XLSetCell(27, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //직위
                    mPrinting.XLSetCell(27, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));  //직군                    
                    mPrinting.XLSetCell(27, 36, iString.ISNull(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6)).Substring(0, 10));  //지급일
                    mPrinting.XLSetCell(28, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));   //입금은행
                    mPrinting.XLSetCell(28, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));  //입금계좌
                    
                    //============================================================================================
                    // 기본급
                    //============================================================================================
                    vAMOUNT = 0; 
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 32, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 32, vAMOUNT);
                    }

                    //============================================================================================
                    // 시급
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 36, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 36, vAMOUNT);
                    }
                    
                    //============================================================================================
                    // 통상 시급
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vAMOUNT  == 0)
                    {
                        mPrinting.XLSetCell(38, 40, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 40, vAMOUNT);
                    }

                    //============================================================================================
                    // 지급합계/공제합계/실지급액
                    //============================================================================================
                    // 월급여 
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_SUPP), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(55, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(55, 15, vAMOUNT);  
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_DED), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(54, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(54, 34, vAMOUNT);   //총공제
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_REAL), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(55, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(55, 34, vAMOUNT);  //실지급액
                    }

                    // 상여액
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_SUPP), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(62, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_DED), 0);
                    if (vAMOUNT == 0)  //총공제액
                    {
                        mPrinting.XLSetCell(61, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 34, vAMOUNT);   //총공제
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_REAL), 0);
                    if (vAMOUNT == 0)  //실지급액
                    {
                        mPrinting.XLSetCell(62, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 34, vAMOUNT);  //실지급액
                    }
                    
                    // 총지급액
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_SUPP), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(63, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(63, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_DED), 0);
                    if (vAMOUNT == 0)  //총공제액
                    {
                        mPrinting.XLSetCell(63, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(63, 34, vAMOUNT);   //총공제
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_REAL), 0);
                    if (vAMOUNT == 0)  //실지급액
                    {
                        mPrinting.XLSetCell(64, 25, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(64, 25, vAMOUNT);  //실지급액
                    }

                    mPrinting.XLSetCell(1, 4, pGrid.GetCellValue(pIndexRow, vIndexWageType));
                    mPrinting.XLSetCell(67, 4, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));  //비고
                }
                else if (pCnt == 2) {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //지급항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //지급액                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(pAllowance_Row+pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row+pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //공제항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //공제액                    

                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("OVER_TIME");        //연장(평일)
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("NIGHT_BONUS_TIME"); //야간(평일)
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LATE_TIME");        //근태공제(평일)
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("HOLY_1_TIME");      //근무(주휴일)
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("HOLY_1_OT");        //연장(주휴일)
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("HOLY_1_NIGHT");     //야간(주휴일)
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("HOLY_0_TIME");      //근무(무휴일)
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("HOLY_0_OT");        //연장(무휴일)
                    int vIndexDataColumn9 = pGrid.GetColumnToIndex("HOLY_0_NIGHT");     //야간(무휴일)
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");   //근무(부가내역)
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("DUTY_30");         //공가(부가내역)
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");  //주차(부가내역)
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("HOLY_1_COUNT");    //유휴(부가내역)
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("HOLY_0_COUNT");    //무휴(부가내역)
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("TOT_DED_COUNT");   //미근무(부가내역)
                    int vIndexDataColumn16 = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");//미주차(부가내역)

                    //============================================================================================
                    // 연장(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 12, "");
                    }
                    else 
                    {
                        mPrinting.XLSetCell(32, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근태공제(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 20, vDUTY_TIME);
                    }                   
                    
                    //============================================================================================
                    // 근무(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 연장(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 12,vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근무(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 연장(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근무(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 4, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 4, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 공가(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 주차(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 유휴(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 무휴(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 미근무(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 24, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 24, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 미주차(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 28, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 28, vDUTY_TIME);
                    }
                }
                else if (pCnt == 5)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //지급항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //지급액                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(56 + pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(56 + pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 6)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //공제항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //공제액                    

                    mPrinting.XLSetCell(56 + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(56 + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }
        
        private void XLContentWrite2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, int pAllowance_Row, int nAllowance_Column)
        {
            decimal vAMOUNT = 0;
            decimal vDUTY_TIME = 0;
            try
            {
                mPrinting.XLActiveSheet("SourceTab2");
                if (pCnt == 1)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("NAME");                   //성명
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("PERSON_NUM");             //사번
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DEPT_NAME");              //부서
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("POST_NAME");              //직위
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("JOB_CLASS_NAME");         //직군
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("SUPPLY_DATE");            //지급일
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("BANK_NAME");              //입금은행
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");          //입금계좌                 
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("BASIC_AMOUNT");          //기본급
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("BASIC_TIME_AMOUNT");     //시급
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("DESCRIPTION");           //비고
                    // '비고'는 후에 추가로 삽입된 것이라 Column 순서가 15로 된 것임.

                    // 명세서 Report 상단에 출력될 내용
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT"); //통상시급
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("WAGE_TYPE");             //급상여구분명
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("PAY_YYYYMM");            //지급년월                    

                    int vIndexWageType = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //지급구분

                    int vIDX_TOT_REAL = pGrid.GetColumnToIndex("REAL_AMOUNT");                // 총 실지급액
                    int vIDX_TOT_SUPP = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          // 총지급액
                    int vIDX_TOT_DED = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");              // 총공제액

                    // 첫장 지급구분/부서/직위/사번/이름
                    mPrinting.XLSetCell(11, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));   //부서
                    mPrinting.XLSetCell(13, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //직위
                    mPrinting.XLSetCell(15, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));   //사번
                    mPrinting.XLSetCell(17, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //이름                    

                    // 인적사항.
                    mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //성명
                    mPrinting.XLSetCell(26, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));  //사번
                    mPrinting.XLSetCell(26, 36, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));  //부서
                    mPrinting.XLSetCell(27, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //직위
                    mPrinting.XLSetCell(27, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));  //직군                    
                    mPrinting.XLSetCell(27, 36, iString.ISNull(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6)).Substring(0, 10));  //지급일
                    mPrinting.XLSetCell(28, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));   //입금은행
                    mPrinting.XLSetCell(28, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));  //입금계좌

                    //============================================================================================
                    // 기본급
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 32, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 32, vAMOUNT);
                    }

                    //============================================================================================
                    // 시급
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 36, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 36, vAMOUNT);
                    }

                    //============================================================================================
                    // 통상 시급
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 40, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 40, vAMOUNT);
                    }

                    //============================================================================================
                    // 지급합계/공제합계/실지급액
                    //============================================================================================
                    // 총지급액
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_SUPP), 0);
                    if (vAMOUNT == 0)  //총지급액
                    {
                        mPrinting.XLSetCell(61, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_DED), 0);
                    if (vAMOUNT == 0)  //총공제액
                    {
                        mPrinting.XLSetCell(61, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 34, vAMOUNT);   //총공제
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_REAL), 0);
                    if (vAMOUNT == 0)  //실지급액
                    {
                        mPrinting.XLSetCell(62, 25, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 25, vAMOUNT);  //실지급액
                    }

                    mPrinting.XLSetCell(1, 4, pGrid.GetCellValue(pIndexRow, vIndexWageType));
                    mPrinting.XLSetCell(67, 4, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));  //비고
                }
                else if (pCnt == 2)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //지급항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //지급액                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //공제항목
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //공제액                    

                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("OVER_TIME");        //연장(평일)
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("NIGHT_BONUS_TIME"); //야간(평일)
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LATE_TIME");        //근태공제(평일)
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("HOLY_1_TIME");      //근무(주휴일)
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("HOLY_1_OT");        //연장(주휴일)
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("HOLY_1_NIGHT");     //야간(주휴일)
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("HOLY_0_TIME");      //근무(무휴일)
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("HOLY_0_OT");        //연장(무휴일)
                    int vIndexDataColumn9 = pGrid.GetColumnToIndex("HOLY_0_NIGHT");     //야간(무휴일)
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");   //근무(부가내역)
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("DUTY_30");         //공가(부가내역)
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");  //주차(부가내역)
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("HOLY_1_COUNT");    //유휴(부가내역)
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("HOLY_0_COUNT");    //무휴(부가내역)
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("TOT_DED_COUNT");   //미근무(부가내역)
                    int vIndexDataColumn16 = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");//미주차(부가내역)

                    //============================================================================================
                    // 연장(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근태공제(평일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근무(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 연장(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(주휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근무(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 연장(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 야간(무휴일)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 근무(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 4, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 4, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 공가(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 주차(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 유휴(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 무휴(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 미근무(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 24, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 24, vDUTY_TIME);
                    }

                    //============================================================================================
                    // 미주차(부가내역)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 28, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 28, vDUTY_TIME);
                    }
                }                
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Excel Wirte Methods ----

        // Excel Wirte Methods 1(급여/상여 인쇄)
        public int XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory
            , string pPeriodFrom, string pUserName, string pCaption, int pCnt)
        {
            string vMessageText = string.Empty;

            int vPageNumber = 0;
            int vTotalRow = pGrid.RowCount; //Grid의 총 행수
            int nAllowance_Row = 42;
            int nAllowance_Column = 6;

            try
            {
                if (pCnt != 1)
                {
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, vTotalRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }
                else if(pCnt == 1)
                {
                    for (int vRow = 0; vRow <= pRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, pRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }

                if (pCnt == 6)
                {
                    //[Sheet2]내용을 [Sheet1]에 붙여넣기
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy, "SourceTab1");
                    XLContentClear();                    
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return vPageNumber;
        }

        // Excel Wirte Methods 2(급여 인쇄)
        public int XLWirte2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory
            , string pPeriodFrom, string pUserName, string pCaption, int pCnt)
        {
            string vMessageText = string.Empty;

            int vPageNumber = 0;
            int vTotalRow = pGrid.RowCount; //Grid의 총 행수
            int nAllowance_Row = 42;
            int nAllowance_Column = 6;

            try
            {
                if (pCnt != 1)
                {
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite2(pGrid, vRow, vTotalRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }
                else if (pCnt == 1)
                {
                    for (int vRow = 0; vRow <= pRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite2(pGrid, vRow, pRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }

                if (pCnt == 4)
                {
                    //[Sheet2]내용을 [Sheet1]에 붙여넣기
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy, "SourceTab2");
                    XLContentClear2();
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return vPageNumber;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(int pCopySumPrintingLine, string pSourceTab)
        {
            int vPrintHeaderColumnSTART = mXLColumnAreaSTART; //복사되어질 쉬트의 폭, 시작열
            int vPrintHeaderColumnEND = mXLColumnAreaEND;     //복사되어질 쉬트의 폭, 종료열

            int vCopySumPrintingLine = 0;
            vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mMaxIncrementCopy;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                //mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLActiveSheet(pSourceTab); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mMaxIncrementCopy, vPrintHeaderColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

                mPrinting.XLPrinting(1, 1);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return 1; // vCopySumPrintingLine;            
            //mPrinting.XLPrintPreview();
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPrinting(pPageSTART, pPageEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        public void PreView()
        {
            try
            {
                mPrinting.XLPrintPreview();
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            try
            {
                System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
                vMaxNumber = vMaxNumber + 1;
                string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

                vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
                mPrinting.XLSave(vSaveFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

    }
}
#endregion;