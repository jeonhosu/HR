using System;
using ISCommonUtil;

namespace HRMF0364
{
    public class XLPrinting
    {
        
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private XL.XLPrint mPrinting = null;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SourceTab1";
        //private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;   //최종 인쇄 라인.
        //private int m1stPrintingLastRow = 0;
        private int mCurrentRow = 0;        //현재 인쇄되는 row 위치.
        //private int mDefaultEndPageRow = 1; // 페이지 증가후 PageCount 기본값.
        private int mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.

        //private string[] mGridColumn; 

        //Copy할때 병합해야할 셀의 행 위치 기억
        private int[] mRowMerge = new int[8] { -1, -1, -1, -1, -1, -1, -1, -1 };
        private int mCountRow = 0; //병합해야할 셀의 행 위치 Count 

        private object mPringingDateTime = string.Empty;
        private object mReq_Person_Name = string.Empty;

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        //public int PrintingLineMAX
        //{
        //    set
        //    {
        //        mPrintingLineMAX = value;
        //    }
        //}

        //public int IncrementCopyMAX
        //{
        //    set
        //    {
        //        mIncrementCopyMAX = value;
        //    }
        //}

        //public int PositionPrintLineSTART
        //{
        //    set
        //    {
        //        mPositionPrintLineSTART = value;
        //    }
        //}

        //public int CopySumPrintingLine
        //{
        //    set
        //    {
        //        mCopySumPrintingLine = value;
        //    }
        //}

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterfaceAdv, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterfaceAdv;
            mMessageAdapter = pMessageAdapter;
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

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
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

                int vCutRight = 2;
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

        #region ----- Line Clear All Methods ----

        private void XlAllLineClear(object pCORP_NAME)
        {
            object vObject = null;

            mPrinting.XLActiveSheet(mTargetSheet);

            int vStartRow = mCurrentRow;
            int vStartCol = mCopy_StartCol + 1;
            int vEndRow = mCopyLineSUM;
            int vEndCol = mCopy_EndCol - 1;

            mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
            mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
            mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
            mPrinting.XL_LineDraw_Top(vStartRow, vStartCol, vEndCol, 2);  //끝에 공백이 있어서.

            mPrinting.XLCellMerge(mCurrentRow, 51, mCurrentRow, vEndCol, true);
            mPrinting.XLCellAlignmentHorizontal(mCurrentRow, 51, mCurrentRow, vEndCol, "R");
            mPrinting.XLSetCell(mCurrentRow, 51, pCORP_NAME);
        }

        private void RateLineClear(int pPrintingLine, int pCopyPrintingRowSTART, int pCopyPrintingRowEnd)
        {

            int vStartRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;
            int vStartCol = mCopy_StartCol + 1;
            int vEndRow = pCopyPrintingRowEnd - 1;
            int vEndCol = mCopy_EndCol;
            int vDrawRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;

            mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
            mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
            mPrinting.XL_LineDraw_Top(vDrawRow, vStartCol, vEndCol, 2);
        }

        #endregion;

        #region ----- Cell Merge Methods ----

        private void CellMerge(int pCopySumPrintingLine, int pCountRow, int[] pRowMerge)
        {            
            //int vXLine = 0;
            int vCountRowMerge = pRowMerge.Length;

            try
            {
                for (int vCount = 0; vCount < vCountRowMerge; vCount++)
                {
                    if (pRowMerge[vCount] == 1)
                    {
                        //vXLine = pCopySumPrintingLine + mPositionPrintLineSTART + (vCount * 4);
                        //int vStartRow = vXLine - 1;
                        //int vStartCol = mXLColumn[1];
                        //int vEndRow = vXLine + 2;
                        //int vEndCol = mXLColumn[3] - 1;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                        //vXLine = pCopySumPrintingLine + mPositionPrintLineSTART + (vCount * 4);
                        //int vStartRow = vXLine - 1;
                        //int vStartCol = mXLColumn[1];
                        //int vEndRow = vXLine;
                        //int vEndCol = mXLColumn[3] - 1;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                        //vStartRow = vXLine + 1;
                        //vEndRow = vXLine + 2;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);
                    }

                    mRowMerge[vCount] = -1;
                }

                mCountRow = 0; //병합해야할 셀의 행 위치 Count, 0으로 Set
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }        
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(object pUserName, object pPrintingDateTime, object pDepartment_NAME)
        { 
            try
            {               
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                
                //출력자 
                if (iConv.ISNull(pUserName) != string.Empty)
                {
                    mPrinting.XLSetCell(34, 1, pUserName);
                } 
                 
                //작업장 
                mPrinting.XLSetCell(2, 24, pDepartment_NAME);
                 
                //출력일자 
                if (iConv.ISNull(pPrintingDateTime) != string.Empty)
                {
                    mPrinting.XLSetCell(34, 53, string.Format("{0:yyyy-MM-dd hh:mm:dd}", pPrintingDateTime));
                }
                else
                {
                    mPrinting.XLSetCell(34, 53, null);
                } 
            }
            catch (System.Exception ex)
            {
                mAppInterface.OnAppMessage(ex.Message);

                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Array Set ----

        private void SetArray(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGridColumn)
        {
            pGridColumn = new int[81];

            pGridColumn[0] = pGrid.GetColumnToIndex("DEPT_NAME");                   //부서
            pGridColumn[1] = pGrid.GetColumnToIndex("POST_NAME");                   //직위
            pGridColumn[2] = pGrid.GetColumnToIndex("PERSON_NUM");                  //사원번호
            pGridColumn[3] = pGrid.GetColumnToIndex("NAME");                        //성명
            pGridColumn[4] = pGrid.GetColumnToIndex("PAY_TYPE_NAME");               //급여제명
            pGridColumn[5] = pGrid.GetColumnToIndex("JOIN_DATE");                   //입사일자.
            pGridColumn[6] = pGrid.GetColumnToIndex("RETIRE_DATE");                 //퇴사일자
            pGridColumn[7] = pGrid.GetColumnToIndex("LONG_YEAR");                   //근속년수
            pGridColumn[8] = pGrid.GetColumnToIndex("LONG_MONTH");                  //근속월수
            pGridColumn[9] = pGrid.GetColumnToIndex("BASIC_BASE_AMOUNT");           //기본급
            pGridColumn[10] = pGrid.GetColumnToIndex("GENERAL_HOURLY_PAY_AMOUNT");  //통상시급

            pGridColumn[11] = pGrid.GetColumnToIndex("PAY_DAY");                    //급여일수
            pGridColumn[12] = pGrid.GetColumnToIndex("A01");                        //기본급
            pGridColumn[13] = pGrid.GetColumnToIndex("LATE_TIME");                  //지각/조퇴
            pGridColumn[14] = pGrid.GetColumnToIndex("A17");                        //근태공제금액
            pGridColumn[15] = pGrid.GetColumnToIndex("OVER_TIME");                  //연장시간
            pGridColumn[16] = pGrid.GetColumnToIndex("A12");                        //연장금액

            pGridColumn[17] = pGrid.GetColumnToIndex("HOLY_1_TIME");                //휴일근로시간
            pGridColumn[18] = pGrid.GetColumnToIndex("A14");                        //휴일근로금액
            pGridColumn[19] = pGrid.GetColumnToIndex("HOLY_0_TIME");                //토요근로시간
            pGridColumn[20] = pGrid.GetColumnToIndex("A20");                        //토요근로금액
            pGridColumn[21] = pGrid.GetColumnToIndex("NIGHT_BONUS");                //야간할증시간
            pGridColumn[22] = pGrid.GetColumnToIndex("A13");                        //야간할증금액

            pGridColumn[23] = pGrid.GetColumnToIndex("A02");                        //직책수당
            pGridColumn[24] = pGrid.GetColumnToIndex("A11");                        //시간외수당
            pGridColumn[25] = pGrid.GetColumnToIndex("A25");                        //차량유지비
            pGridColumn[26] = pGrid.GetColumnToIndex("A30");                        //교통비
            pGridColumn[27] = pGrid.GetColumnToIndex("A32");                        //토요근무수당
            pGridColumn[28] = pGrid.GetColumnToIndex("A22");                        //결근공제
            pGridColumn[29] = pGrid.GetColumnToIndex("A24");                        //년차수당
            pGridColumn[30] = pGrid.GetColumnToIndex("A09");                        //상여금
            pGridColumn[31] = pGrid.GetColumnToIndex("A07");                        //기타수당
            pGridColumn[32] = pGrid.GetColumnToIndex("A28");                        //만근수당
            pGridColumn[33] = pGrid.GetColumnToIndex("A27");                        //철야수당
            pGridColumn[34] = pGrid.GetColumnToIndex("A37");                        //
            pGridColumn[35] = pGrid.GetColumnToIndex("A38");                        //
            pGridColumn[36] = pGrid.GetColumnToIndex("A37");                        //
            pGridColumn[37] = pGrid.GetColumnToIndex("A39");                        //
            pGridColumn[38] = pGrid.GetColumnToIndex("A38");                        //
            pGridColumn[39] = pGrid.GetColumnToIndex("ETC_SUM");                    //기타수당합계
            pGridColumn[40] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          //지급총합계

            pGridColumn[41] = pGrid.GetColumnToIndex("D01");                        //소득세
            pGridColumn[42] = pGrid.GetColumnToIndex("D02");                        //주민세            
            pGridColumn[43] = pGrid.GetColumnToIndex("D03");                        //국민연금
            pGridColumn[44] = pGrid.GetColumnToIndex("D04");                        //고용보험
            pGridColumn[45] = pGrid.GetColumnToIndex("D05");                        //건강보험
            pGridColumn[46] = pGrid.GetColumnToIndex("D06");                        //장기요양보험
            pGridColumn[47] = pGrid.GetColumnToIndex("D07");                        //건강보험정산액
            pGridColumn[48] = pGrid.GetColumnToIndex("D08");                        //요양보험정산액
            pGridColumn[49] = pGrid.GetColumnToIndex("D09");                        //가불금
            pGridColumn[50] = pGrid.GetColumnToIndex("D10");                        //전월정산액
            pGridColumn[51] = pGrid.GetColumnToIndex("D11");                        //피복비
            pGridColumn[52] = pGrid.GetColumnToIndex("D12");                        //사원증발급비
            pGridColumn[53] = pGrid.GetColumnToIndex("D13");                        //개인신용보험
            pGridColumn[54] = pGrid.GetColumnToIndex("D14");                        //기타공제
            pGridColumn[55] = pGrid.GetColumnToIndex("D15");                        //정산소득세
            pGridColumn[56] = pGrid.GetColumnToIndex("D16");                        //정산주민세
            pGridColumn[57] = pGrid.GetColumnToIndex("D17");                        //정산농특세
            pGridColumn[58] = pGrid.GetColumnToIndex("D18");                        // 
            pGridColumn[59] = pGrid.GetColumnToIndex("D19");                        //가압류공제 
            pGridColumn[60] = pGrid.GetColumnToIndex("D20");                        // 

            pGridColumn[61] = pGrid.GetColumnToIndex("D21");                        // 
            pGridColumn[62] = pGrid.GetColumnToIndex("D22");                        // 
            pGridColumn[63] = pGrid.GetColumnToIndex("D23");                        // 
            pGridColumn[64] = pGrid.GetColumnToIndex("D24");                        // 
            pGridColumn[65] = pGrid.GetColumnToIndex("D25");                        //연말정산소득세
            pGridColumn[66] = pGrid.GetColumnToIndex("D26");                        //연말정산주민세
            pGridColumn[67] = pGrid.GetColumnToIndex("D27");                        //연말정산농특세
            pGridColumn[68] = pGrid.GetColumnToIndex("D28");                        //상조회비
            pGridColumn[69] = pGrid.GetColumnToIndex("D29");                        //  

            pGridColumn[70] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");             //총공제액 
            pGridColumn[71] = pGrid.GetColumnToIndex("REAL_AMOUNT");                //실지급액

            pGridColumn[72] = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");              //출근일주
            pGridColumn[73] = pGrid.GetColumnToIndex("DUTY_30");                    //공가
            pGridColumn[74] = pGrid.GetColumnToIndex("TOT_DED_COUNT");              //미근무
            pGridColumn[75] = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");             //주차
            pGridColumn[76] = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");           //미주차
            pGridColumn[77] = pGrid.GetColumnToIndex("HOLY_1_COUNT");               //유휴
            pGridColumn[78] = pGrid.GetColumnToIndex("HOLY_0_COUNT");               //무휴
            pGridColumn[79] = pGrid.GetColumnToIndex("DEPT_CODE");                  //부서코드
            pGridColumn[80] = pGrid.GetColumnToIndex("SUMMARY_FLAG");               //합계여부
        }

        #endregion;

        #region ----- Convert String Methods ----

        private string ConvertString(object pObject)
        {
            string vString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is string;
                    if (IsConvert == true)
                    {
                        vString = pObject as string;
                    }
                }
            }
            catch
            {
            }

            return vString;
        }

        #endregion;

        #region ----- Convert DateTime Methods ----

        private string ConvertDateTime(object pObject)
        {
            string vTextDateTimeLong = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vTextDateTimeLong;
        }

        private string ConvertDate(object pObject)
        {
            string vTextDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        vTextDateTimeShort = vDateTime.ToShortDateString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vTextDateTimeShort;
        }

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {
            bool vIsConvert = false;
            pConvertString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        pConvertString = pObject as string;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        pConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(string pStringNumber, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pStringNumber != null)
                {
                    decimal vIsConvertNum = decimal.Parse(pStringNumber);
                    pConvertDecimal = vIsConvertNum;
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out string pConvertDateTimeShort)
        {
            bool vIsConvert = false;
            pConvertDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime.ToShortDateString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- Line Print ----- 
         
        private int XlLine(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vGetValue = null;  

            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            //string vSUMMARY_FLAG = "N";

            bool IsConvert = false;  
            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //[성명] 
                vGetValue = pRow["NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {
                    
                }
                else
                {
                    vConvertString = string.Empty;
                } 
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //[직위] 
                vGetValue = pRow["POST_NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty; 
                }
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //[직구분] 
                vGetValue = pRow["JOB_CATEGORY_NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vConvertString);

                //[근무일자] 
                vGetValue = pRow["WORK_DATE"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 13, vConvertString);

                //[당직] 
                vGetValue = pRow["DANGJIK_YN"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //[철야] 
                vGetValue = pRow["ALL_NIGHT_YN"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vConvertString);

                //[근무전-시작] 
                vGetValue = pRow["BEFORE_TIME_START"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vConvertString);

                //[근무전-종료] 
                vGetValue = pRow["BEFORE_TIME_END"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //[근무후-시작] 
                vGetValue = pRow["AFTER_OT_START"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vConvertString);

                //[근무전-종료] 
                vGetValue = pRow["AFTER_OT_END"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vConvertString);

                //[조식] 
                vGetValue = pRow["BREAKFAST_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 42, vConvertString);

                //[중식] 
                vGetValue = pRow["LUNCH_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vConvertString);

                //[석식] 
                vGetValue = pRow["DINNER_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 46, vConvertString);

                //[야식] 
                vGetValue = pRow["MIDNIGHT_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 48, vConvertString);

                //[사유]
                vGetValue = pRow["DESCRIPTION"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 50, vConvertString); 

                vXLine = vXLine + 1;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            } 
            return vXLine;
        }

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int XLWirteMain(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter
                                , object pLocal_DATE
                                , object pReq_Person_Name)
        {
            string vMessage = string.Empty;
            mIsNewPage = false; 

            //초기화//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 62;
            mCopy_EndRow = 34;

            //mDefaultEndPageRow = 1;
            mDefaultPageRow = 8;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 33;  //최종 인쇄 라인.
            //m1stPrintingLastRow = 40;

            mCurrentRow = 8;
            mCopyLineSUM = 1;

            int vTotalRow = 0;
            int vPageRowCount = 0;  //인쇄후 해당 라인 증가 위해. 
            int vCurrRow = 0;

            mPringingDateTime = pLocal_DATE;

            string vDEPT_CODE = string.Empty; 
            try
            {
                vTotalRow = pAdapter.CurrentRows.Count;
                //TotalPage(pGrid);

                if (vTotalRow > 0)
                {
                    //배열 정의. 
                    vPageRowCount = mCurrentRow - 1;  
                    foreach(System.Data.DataRow vRow in pAdapter.CurrentRows)
                    {
                        vCurrRow++;
                        vMessage = string.Format("Row : {0} / {1}", vCurrRow, vTotalRow);
                        mAppInterface.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        if (vCurrRow == 1)
                        {
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pReq_Person_Name, vRow["FLOOR_NAME"]);
                        }
                        //if (vRow == 0)
                        //{
                        //    //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pGrid, vRow, vDEPT_NAME);
                        //    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME); 
                        //}
                        //else if (vDEPT_CODE != iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[79])) && mIsNewPage == false)
                        //{
                        //    //XlAllLineClear(pCorporationName);
                        //    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME);
                        //    //아직인쇄 전 이므로 페이지ROW에 +4를 해줌.
                        //    mCurrentRow = mCurrentRow + (mCopy_EndRow - (vPageRowCount + 4)) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                        //    vPageRowCount = mDefaultPageRow - 4;
                        //}

                        mCurrentRow = XlLine(vRow, mCurrentRow);
                        vPageRowCount = vPageRowCount + 1;
                        
                        IsNewPage(mPrinting, vPageRowCount, vDEPT_CODE, iConv.ISNull(vRow["FLOOR_CODE"]), pReq_Person_Name, vRow["FLOOR_NAME"]);   // 새로운 페이지 체크 및 생성.
                        if (mIsNewPage == true)
                        {
                            //인쇄 후 이므로 현재 페이지ROW에 -4를 해줌.
                            mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 1) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                            vPageRowCount = mDefaultPageRow - 1;
                        }
                        vDEPT_CODE = iConv.ISNull(vRow["FLOOR_CODE"]);

                        //if (vRow == vTotalRow -1)
                        //{
                        //    // 마지막 데이터 이면 처리할 사항 기술
                        //    // 라인지운다 또는 합계를 표시한다 등 기술.
                        //    SumWrite(mCurrentRow);      //합계.
                        //    if (vPageRowCount != mPrintingLastRow)
                        //    {
                        //        //마지막ROW가 마지막 인쇄하고 다르면 엑셀 라인 CLEAR
                        //        XlAllLineClear(pCorporationName);
                        //    }
                        //}
                        //else
                        //{
                        //    IsNewPage(vPageRowCount, false, vDEPT_NAME);   // 새로운 페이지 체크 및 생성.
                        //    if (mIsNewPage == true)
                        //    {
                        //        //인쇄 후 이므로 현재 페이지ROW에 -4를 해줌.
                        //        mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 4) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                        //        vPageRowCount = mDefaultPageRow - 4;
                        //    }
                        //} 
                    }

                    //mPrinting.XLDeleteSheet(mSourceSheet1); 
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        private void SumWrite(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageCount = 45;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageCount * (r - 1);
                mPrinting.XLSetCell((vLINE + 4), 56, string.Format("Page {0} of {1}", r, mPageNumber));

                //if (r == mPageNumber)
                //{
                //    //
                //}
                //else
                //{
                //    vLINE = vLINE - 1;
                //    mPrinting.XLSetCell(vLINE, 1, "");
                //}
            }

            ////합계 인쇄//
            //vLINE = mPageNumber * mCopy_EndRow;
            //vLINE = vLINE - 1;
            ////mPrinting.XLSetCell(vLINE, 1, "SUM");
            //string vAmount = string.Empty;

            ////[합계]
            //if (mPageNumber == 1)
            //{
            //    vLINE = 31;
            //    mPrinting.XLSetCell(vLINE, 1, "[총    계]");

            //    //BACK COLOR.
            //    mPrinting.XLCellColorBrush(vLINE, 8, vLINE, 15, System.Drawing.Color.Silver);

            //    //계획합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 8, vAmount);

            //    //예산합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 11, vAmount);

            //    //차액합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 14, vAmount);

            //    //XlLineClear(pPrintingLine);

            //}
            //else
            //{
            //    mPrinting.XLSetCell(vLINE, 1, "[총    계]");

            //    //BACK COLOR.
            //    mPrinting.XLCellColorBrush(vLINE, 8, vLINE, 15, System.Drawing.Color.Silver);

            //    //계획합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 8, vAmount);

            //    //예산합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 11, vAmount);

            //    //차액합계
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 14, vAmount);

            //    //XlLineClear(pPrintingLine);
            //}
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, object pDEPT_CODE, object pReq_Person_Name, object pDEPT_NAME)
        {
            if (mPrintingLastRow == pPrintingLine)
            {
                mIsNewPage = true;                
                //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true; 
                //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pPrintingLine, pGrid, pRow);  
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pPrintingLine, pGrid, pRow); 
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage(XL.XLPrint pPrinting, int pPrintingLine, object pOLD_DEPT_CODE, object pDEPT_CODE, object pReq_Person_Name, object pDEPT_NAME)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(pPrinting, mCopyLineSUM, pReq_Person_Name, pDEPT_NAME);
            }
            else if (iConv.ISNull(pOLD_DEPT_CODE) != string.Empty && iConv.ISNull(pOLD_DEPT_CODE) != iConv.ISNull(pDEPT_CODE))
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(pPrinting, mCopyLineSUM, pReq_Person_Name, pDEPT_NAME); 
            } 
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, object pReq_Person_Name, object pDEPT_NAME)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            HeaderWrite(pReq_Person_Name, mPringingDateTime, pDEPT_NAME);
            //DepartmentName(pGrid, pRow);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;

            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, object pDEPT_NAME)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            //HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);            
            //DepartmentName(pGrid, pRow);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;

            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(int pCopySumPrintingLine, int pPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        {
            int vPrintHeaderColumnSTART = mCopy_StartCol; //복사되어질 쉬트의 폭, 시작열
            int vPrintHeaderColumnEND = mCopy_EndCol;     //복사되어질 쉬트의 폭, 종료열

            mPageNumber++;
            //mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
            //HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mCorporationName);
            //DepartmentName(pGrid, pRow);

            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            mPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mCopy_EndRow, vPrintHeaderColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            //업체
            int vDrawRow = (pPrintingLine + vCopyPrintingRowSTART) - 1;
            //mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

            CellMerge(pCopySumPrintingLine, mCountRow, mRowMerge);

            RateLineClear(pPrintingLine, vCopyPrintingRowSTART, vCopyPrintingRowEnd);

            return vCopySumPrintingLine;
        }

        ////[Sheet2]내용을 [Sheet1]에 붙여넣기
        //private int CopyAndPaste_1(int pCopySumPrintingLine, int pPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        //{
        //    int vPrintHeaderColumnSTART = mCopy_StartCol; //복사되어질 쉬트의 폭, 시작열
        //    int vPrintHeaderColumnEND = mCopy_EndCol;     //복사되어질 쉬트의 폭, 종료열

        //    mCountPage++;
        //    mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
        //    HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mPageString, mCorporationName);
        //    DepartmentName(pGrid, pRow);

        //    int vCopySumPrintingLine = pCopySumPrintingLine;

        //    int vCopyPrintingRowSTART = vCopySumPrintingLine;
        //    vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
        //    int vCopyPrintingRowEnd = vCopySumPrintingLine;
        //    mPrinting.XLActiveSheet("SourceTab1");
        //    object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mIncrementCopyMAX, vPrintHeaderColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    mPrinting.XLActiveSheet("Destination");
        //    object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

        //    //업체
        //    int vDrawRow = (pPrintingLine + vCopyPrintingRowSTART) - 1;
        //    mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

        //    CellMerge(pCopySumPrintingLine, mCountRow, mRowMerge);

        //    RateLineClear(pPrintingLine, vCopyPrintingRowSTART, vCopyPrintingRowEnd);

        //    return vCopySumPrintingLine;
        //}

        #endregion;

        //#region ----- Excel Rate Line Clear Method ----

        //private void RateLineClear(int pPrintingLine, int pCopyPrintingRowSTART, int pCopyPrintingRowEnd)
        //{

        //    int vStartRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;
        //    int vStartCol = mCopyColumnSTART + 1;
        //    int vEndRow = pCopyPrintingRowEnd - 1;
        //    int vEndCol = mCopyColumnEND;
        //    int vDrawRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;

        //    mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
        //    mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
        //    mPrinting.XL_LineDraw_Top(vDrawRow, vStartCol, vEndCol, 2);
        //}

        //#endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPrinting(pPageSTART, pPageEND);
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }
            mPrinting.XLSave(pSaveFileName);

            //전호수 주석
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- Save Methods ----

        public void PDF_Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }
            mPrinting.XLSaveAs_PDF(pSaveFileName);

            //전호수 주석
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.

            int vXLRow = 31; //엑셀에 내용이 표시되는 행 번호
            int vXLCol = 40;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(pActiveSheet);
                mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;
                
    }
}