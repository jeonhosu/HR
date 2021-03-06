using System;
using ISCommonUtil;

namespace HRMF0507
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
        private string mTargetSheet = "Destination";
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

        private object mCorporationName = string.Empty;
        private object mUserName = string.Empty;
        private object mYYYYMM = string.Empty;
        private object mWageTypeName = string.Empty;
        private object mDepartmentName = string.Empty;
        private object mPringingDateTime = string.Empty;

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
            int vStartCol = mCopy_StartCol;
            int vEndRow = mCopyLineSUM;
            int vEndCol = mCopy_EndCol;

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

        public void HeaderWrite(object pUserName, object pPrintingDateTime, object pYYYYMM, object pWageTypeName, object pDepartment_NAME, object pCorporationName)
        { 
            try
            {
                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(2, 1);    //Title
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(4, 6);    //출력자
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(5, 6);    //급여구분
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 20);   //부서
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(4, 58);   //페이지
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(5, 58);   //출력일자
                System.Drawing.Point vCellPoint07 = new System.Drawing.Point(48, 53);  //업체

                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택

                //Title 
                if (iConv.ISNull(pYYYYMM) != string.Empty)
                {
                    string vYear = iConv.ISNull(pYYYYMM).Substring(0, 4);
                    string vMonth = iConv.ISNull(pYYYYMM).Substring(5, 2);
                    string vTitle = string.Format("{0}년 {1}월 {2} 대장", vYear, vMonth, pWageTypeName);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vTitle);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, null);
                }

                //출력자 
                if (iConv.ISNull(pUserName) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, pUserName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, null);
                }

                //급여구분 
                if (iConv.ISNull(pWageTypeName) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, pWageTypeName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, "전체");
                }

                //부서 
                mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, pDepartment_NAME);

                ////페이지 
                //if (iConv.ISNull(pPageString) != string.Empty)
                //{
                //    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, pPageString);
                //}
                //else
                //{
                //    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, null);
                //}

                //출력일자 
                if (iConv.ISNull(pPrintingDateTime) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, string.Format("{0:yyyy-MM-dd hh:mm:dd}", pPrintingDateTime));
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, null);
                }

                //업체
                mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, pCorporationName); 
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
            pGridColumn = new int[83];

            pGridColumn[0] = pGrid.GetColumnToIndex("DEPT_NAME");                   //부서
            pGridColumn[1] = pGrid.GetColumnToIndex("POST_NAME");                   //직위
            pGridColumn[2] = pGrid.GetColumnToIndex("PERSON_NUM");                  //사원번호
            pGridColumn[3] = pGrid.GetColumnToIndex("NAME");                        //성명
            pGridColumn[4] = pGrid.GetColumnToIndex("PAY_TYPE_NAME");               //급여제명
            pGridColumn[5] = pGrid.GetColumnToIndex("JOIN_DATE");                   //입사일자.
            pGridColumn[6] = pGrid.GetColumnToIndex("RETIRE_DATE");                 //퇴사일자
            pGridColumn[7] = pGrid.GetColumnToIndex("LONG_YEAR");                   //근속년수
            pGridColumn[8] = pGrid.GetColumnToIndex("LONG_MONTH");                  //근속월수
            pGridColumn[9] = pGrid.GetColumnToIndex("BAISC_AMOUNT");             //월정급
            pGridColumn[10] = pGrid.GetColumnToIndex("GENERAL_HOURLY_PAY_AMOUNT");  //통상시급

            pGridColumn[11] = pGrid.GetColumnToIndex("PAY_DAY");                    //급여일수
            pGridColumn[12] = pGrid.GetColumnToIndex("BASE_A01");                        //기본급
            pGridColumn[13] = pGrid.GetColumnToIndex("LATE_TIME");                  //지각/조퇴
            pGridColumn[14] = pGrid.GetColumnToIndex("A17");                        //지각외출조퇴 근태공제
            pGridColumn[15] = pGrid.GetColumnToIndex("OVER_TIME");                  //연장시간
            pGridColumn[16] = pGrid.GetColumnToIndex("A12");                        //연장근로수당 + 연장수당

            pGridColumn[17] = pGrid.GetColumnToIndex("HOLY_1_TIME");                //휴일근로시간
            pGridColumn[18] = pGrid.GetColumnToIndex("A14");                        //휴일근로금액
            pGridColumn[19] = pGrid.GetColumnToIndex("HOLY_0_TIME");                //휴일연장
            pGridColumn[20] = pGrid.GetColumnToIndex("A13");                        //휴일연장금액
            pGridColumn[21] = pGrid.GetColumnToIndex("NIGHT_BONUS");                //야간할증시간
            pGridColumn[22] = pGrid.GetColumnToIndex("A20");                       //심야수당

            pGridColumn[23] = pGrid.GetColumnToIndex("A02");                        //연장수당(관리직) 연장근로계산금액에 들어감
            pGridColumn[24] = pGrid.GetColumnToIndex("A04");                        //특별수당
            pGridColumn[25] = pGrid.GetColumnToIndex("A05");                        //직책수당
            pGridColumn[26] = pGrid.GetColumnToIndex("A22");                        //기타수당2 생산
            pGridColumn[27] = pGrid.GetColumnToIndex("ETC_SUM");                    //그외수당 
            pGridColumn[28] = pGrid.GetColumnToIndex("A33");                        //기타수당2 관리
            pGridColumn[29] = pGrid.GetColumnToIndex("A15");                        //당직비
            pGridColumn[30] = pGrid.GetColumnToIndex("A24");                        //연차수당
            pGridColumn[31] = pGrid.GetColumnToIndex("A25");                        //차량보조
            pGridColumn[32] = pGrid.GetColumnToIndex("A09");                        //상여금
            pGridColumn[33] = pGrid.GetColumnToIndex("A11");                        //급여소급분
            pGridColumn[34] = pGrid.GetColumnToIndex("A07");                        //기타수당
            pGridColumn[35] = pGrid.GetColumnToIndex("A31");                        //기타수당.포상금.
            pGridColumn[36] = pGrid.GetColumnToIndex("A10");                        // X
            pGridColumn[37] = pGrid.GetColumnToIndex("A03");                        //X
            pGridColumn[38] = pGrid.GetColumnToIndex("A06");                        // X
            pGridColumn[39] = pGrid.GetColumnToIndex("A08");                       //X
            pGridColumn[81] = pGrid.GetColumnToIndex("A24");                       //X
            pGridColumn[82] = pGrid.GetColumnToIndex("A16");                       //X
            pGridColumn[40] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          //지급총합계

            pGridColumn[41] = pGrid.GetColumnToIndex("D01");                        //소득세 //
            pGridColumn[42] = pGrid.GetColumnToIndex("D02");                        //주민세           // 
            pGridColumn[43] = pGrid.GetColumnToIndex("D03");                        //국민연금//
            pGridColumn[44] = pGrid.GetColumnToIndex("D04");                        //고용보험//
            pGridColumn[45] = pGrid.GetColumnToIndex("D05");                        //건강보험//
            pGridColumn[46] = pGrid.GetColumnToIndex("D14");                        //기타공제//
            pGridColumn[47] = pGrid.GetColumnToIndex("D07");                        //건강보험정산액  //
            pGridColumn[48] = pGrid.GetColumnToIndex("D08");                        //요양보험정산액//
            pGridColumn[49] = pGrid.GetColumnToIndex("D22");                        //이자납입

            pGridColumn[50] = pGrid.GetColumnToIndex("D10");                        //전월정산액 //
            pGridColumn[51] = pGrid.GetColumnToIndex("D11");                        //건강보험 연말정산
            pGridColumn[52] = pGrid.GetColumnToIndex("D12");                        //작업복//
            pGridColumn[53] = pGrid.GetColumnToIndex("D13");                        //건강검진//

            pGridColumn[54] = pGrid.GetColumnToIndex("D15");                        //정산소득세//

            pGridColumn[55] = pGrid.GetColumnToIndex("D32");                        //학자금공제
            pGridColumn[56] = pGrid.GetColumnToIndex("D16");                        //정산주민세 //
            pGridColumn[57] = pGrid.GetColumnToIndex("D17");                        //정산농특세//
            pGridColumn[58] = pGrid.GetColumnToIndex("D28");                        // 경조금//
            pGridColumn[59] = pGrid.GetColumnToIndex("D19");                        //건강보험 퇴직정산 // 
            pGridColumn[60] = pGrid.GetColumnToIndex("D20");                        // 산재보험

            pGridColumn[61] = pGrid.GetColumnToIndex("D06");                        // 장기요양보험 // 
            pGridColumn[62] = pGrid.GetColumnToIndex("D23");                        // 가압류//
            pGridColumn[63] = pGrid.GetColumnToIndex("D09");                             //국민연금소급분 
            pGridColumn[64] = pGrid.GetColumnToIndex("");                             // 
            pGridColumn[65] = pGrid.GetColumnToIndex("D25");                        // 연말소득세
            pGridColumn[66] = pGrid.GetColumnToIndex("D26");                        // 연말주민세
            pGridColumn[67] = pGrid.GetColumnToIndex("D27");                        // 연말농특세 
            pGridColumn[68] = pGrid.GetColumnToIndex("D29");                        // 식대공제
            pGridColumn[69] = pGrid.GetColumnToIndex(" ");                            //  

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


        private int XlPrompt(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            int vXLine = 8; //엑셀에 내용이 표시되는 행 번호
            object vValue = string.Empty;

            try
            {  
                mPrinting.XLActiveSheet(mSourceSheet1);

                foreach(System.Data.DataRow vRow in pAdapter.CurrentRows)
                {
                    //1.인적사항. 
                    mPrinting.XLSetCell(vXLine, 1, vRow["P01"]);
                    mPrinting.XLSetCell(vXLine, 5, vRow["P02"]);  
                    //근무내역. 
                    mPrinting.XLSetCell(vXLine, 9, vRow["P03"]);
                    mPrinting.XLSetCell(vXLine, 12, vRow["P04"]);
                    mPrinting.XLSetCell(vXLine, 15, vRow["P05"]);
                    mPrinting.XLSetCell(vXLine, 18, vRow["P06"]); 

                    //지급. 
                    mPrinting.XLSetCell(vXLine, 21, vRow["P07"]);
                    mPrinting.XLSetCell(vXLine, 25, vRow["P08"]);
                    mPrinting.XLSetCell(vXLine, 29, vRow["P09"]);
                    mPrinting.XLSetCell(vXLine, 33, vRow["P10"]);
                    mPrinting.XLSetCell(vXLine, 37, vRow["P11"]); 

                    //공제. 
                    mPrinting.XLSetCell(vXLine, 41, vRow["P12"]);
                    mPrinting.XLSetCell(vXLine, 45, vRow["P13"]);
                    mPrinting.XLSetCell(vXLine, 49, vRow["P14"]);
                    mPrinting.XLSetCell(vXLine, 53, vRow["P15"]);
                    mPrinting.XLSetCell(vXLine, 57, vRow["P16"]); 

                    //합계
                    mPrinting.XLSetCell(vXLine, 61, vRow["P17"]);

                    //2.인적사항. 
                    mPrinting.XLSetCell(vXLine + 1, 1, vRow["P18"]);
                    mPrinting.XLSetCell(vXLine + 1, 5, vRow["P19"]);
                    //근무내역. 
                    mPrinting.XLSetCell(vXLine + 1, 9, vRow["P20"]);
                    mPrinting.XLSetCell(vXLine + 1, 12, vRow["P21"]);
                    mPrinting.XLSetCell(vXLine + 1, 15, vRow["P22"]);
                    mPrinting.XLSetCell(vXLine + 1, 18, vRow["P23"]);

                    //지급. 
                    mPrinting.XLSetCell(vXLine + 1, 21, vRow["P24"]);
                    mPrinting.XLSetCell(vXLine + 1, 25, vRow["P25"]);
                    mPrinting.XLSetCell(vXLine + 1, 29, vRow["P26"]);
                    mPrinting.XLSetCell(vXLine + 1, 33, vRow["P27"]);
                    mPrinting.XLSetCell(vXLine + 1, 37, vRow["P28"]);

                    //공제. 
                    mPrinting.XLSetCell(vXLine + 1, 41, vRow["P29"]);
                    mPrinting.XLSetCell(vXLine + 1, 45, vRow["P30"]);
                    mPrinting.XLSetCell(vXLine + 1, 49, vRow["P31"]);
                    mPrinting.XLSetCell(vXLine + 1, 53, vRow["P32"]);
                    mPrinting.XLSetCell(vXLine + 1, 57, vRow["P33"]);

                    //합계
                    mPrinting.XLSetCell(vXLine + 1, 61, vRow["P34"]);

                    //3.인적사항. 
                    mPrinting.XLSetCell(vXLine + 2, 1, vRow["P35"]);
                    mPrinting.XLSetCell(vXLine + 2, 5, vRow["P36"]);
                    //근무내역. 
                    mPrinting.XLSetCell(vXLine + 2, 9, vRow["P37"]);
                    mPrinting.XLSetCell(vXLine + 2, 12, vRow["P38"]);
                    mPrinting.XLSetCell(vXLine + 2, 15, vRow["P39"]);
                    mPrinting.XLSetCell(vXLine + 2, 18, vRow["P40"]);

                    //지급. 
                    mPrinting.XLSetCell(vXLine + 2, 21, vRow["P41"]);
                    mPrinting.XLSetCell(vXLine + 2, 25, vRow["P42"]);
                    mPrinting.XLSetCell(vXLine + 2, 29, vRow["P43"]);
                    mPrinting.XLSetCell(vXLine + 2, 33, vRow["P44"]);
                    mPrinting.XLSetCell(vXLine + 2, 37, vRow["P45"]);

                    //공제. 
                    mPrinting.XLSetCell(vXLine + 2, 41, vRow["P46"]);
                    mPrinting.XLSetCell(vXLine + 2, 45, vRow["P47"]);
                    mPrinting.XLSetCell(vXLine + 2, 49, vRow["P48"]);
                    mPrinting.XLSetCell(vXLine + 2, 53, vRow["P49"]);
                    mPrinting.XLSetCell(vXLine + 2, 57, vRow["P50"]);

                    //합계
                    mPrinting.XLSetCell(vXLine + 2, 61, vRow["P51"]);

                    //4.인적사항. 
                    mPrinting.XLSetCell(vXLine + 3, 1, vRow["P52"]);
                    mPrinting.XLSetCell(vXLine + 3, 5, vRow["P53"]);
                    //근무내역. 
                    mPrinting.XLSetCell(vXLine + 3, 9, vRow["P54"]);
                    mPrinting.XLSetCell(vXLine + 3, 12, vRow["P55"]);
                    mPrinting.XLSetCell(vXLine + 3, 15, vRow["P56"]);
                    mPrinting.XLSetCell(vXLine + 3, 18, vRow["P57"]);

                    //지급. 
                    mPrinting.XLSetCell(vXLine + 3, 21, vRow["P58"]);
                    mPrinting.XLSetCell(vXLine + 3, 25, vRow["P59"]);
                    mPrinting.XLSetCell(vXLine + 3, 29, vRow["P60"]);
                    mPrinting.XLSetCell(vXLine + 3, 33, vRow["P61"]);
                    mPrinting.XLSetCell(vXLine + 3, 37, vRow["P62"]);

                    //공제. 
                    mPrinting.XLSetCell(vXLine + 3, 41, vRow["P63"]);
                    mPrinting.XLSetCell(vXLine + 3, 45, vRow["P64"]);
                    mPrinting.XLSetCell(vXLine + 3, 49, vRow["P65"]);
                    mPrinting.XLSetCell(vXLine + 3, 53, vRow["P66"]);
                    mPrinting.XLSetCell(vXLine + 3, 57, vRow["P67"]);

                    //합계
                    mPrinting.XLSetCell(vXLine + 3, 61, vRow["P68"]);

                    //5.인적사항. 
                    mPrinting.XLSetCell(vXLine + 4, 1, vRow["P69"]);
                    mPrinting.XLSetCell(vXLine + 4, 5, vRow["P70"]);
                    //근무내역. 
                    mPrinting.XLSetCell(vXLine + 4, 9, vRow["P71"]);
                    mPrinting.XLSetCell(vXLine + 4, 12, vRow["P72"]);
                    mPrinting.XLSetCell(vXLine + 4, 15, vRow["P73"]);
                    mPrinting.XLSetCell(vXLine + 4, 18, vRow["P74"]);

                    //지급. 
                    mPrinting.XLSetCell(vXLine + 4, 21, vRow["P75"]);
                    mPrinting.XLSetCell(vXLine + 4, 25, vRow["P76"]);
                    mPrinting.XLSetCell(vXLine + 4, 29, vRow["P77"]);
                    mPrinting.XLSetCell(vXLine + 4, 33, vRow["P78"]);
                    mPrinting.XLSetCell(vXLine + 4, 37, vRow["P79"]);

                    //공제. 
                    mPrinting.XLSetCell(vXLine + 4, 41, vRow["P80"]);
                    mPrinting.XLSetCell(vXLine + 4, 45, vRow["P81"]);
                    mPrinting.XLSetCell(vXLine + 4, 49, vRow["P82"]);
                    mPrinting.XLSetCell(vXLine + 4, 53, vRow["P83"]);
                    mPrinting.XLSetCell(vXLine + 4, 57, vRow["P84"]);

                    //합계
                    mPrinting.XLSetCell(vXLine + 4, 61, vRow["P85"]);

                }  
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            }
            return vXLine;
        }


        private int XlLine(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vGetValue = null;
            object vGetValue_2 = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            decimal vConvertDecimal2 = 0m;

            decimal vTEMP_AMT = 0m;
            string vSUMMARY_FLAG = "N";

            bool IsConvert = false;
            bool IsConvert2 = false;
            try
            {
                vSUMMARY_FLAG = iConv.ISNull(pRow["SUMMARY_FLAG"]);

                mPrinting.XLActiveSheet(mTargetSheet);

                //1번 부서명 출력.
                vGetValue = pRow["P01"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {

                }
                else
                {
                    mPrinting.XLCellMerge(vXLine, 1, vXLine + 1, 8, true);
                }
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //2번
                vGetValue = pRow["P02"]; 
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                }

                //3번
                vGetValue = pRow["P03"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //4번
                vGetValue = pRow["P04"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //5번
                vGetValue = pRow["P05"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //6번
                vGetValue = pRow["P06"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //7번
                vGetValue = pRow["P07"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vConvertString);

                //8번
                vGetValue = pRow["P08"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //9번
                vGetValue = pRow["P09"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //10번
                vGetValue = pRow["P10"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

                //11
                vGetValue = pRow["P11"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vConvertString);

                //12
                vGetValue = pRow["P12"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vConvertString);

                //13
                vGetValue = pRow["P13"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vConvertString);

                //14
                vGetValue = pRow["P14"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //15
                vGetValue = pRow["P15"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 53, vConvertString);

                //16
                vGetValue = pRow["P16"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 57, vConvertString);

                //17
                vGetValue = pRow["P17"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 61, vConvertString);


                //////////////////////////////////////////////////////////////////////////////////////////
                ///2번째 라인///                
                vXLine++; 
                ///////////////
                vGetValue = pRow["P18"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 1, vConvertString);
                }

                //19
                vGetValue = pRow["P19"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                }

                //20번
                vGetValue = pRow["P20"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //21번
                vGetValue = pRow["P21"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //22번
                vGetValue = pRow["P22"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //23번
                vGetValue = pRow["P23"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //24번
                vGetValue = pRow["P24"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vConvertString);

                //25번
                vGetValue = pRow["P25"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //26번
                vGetValue = pRow["P26"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //27번
                vGetValue = pRow["P27"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

                //28
                vGetValue = pRow["P28"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vConvertString);

                //29
                vGetValue = pRow["P29"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vConvertString);

                //30
                vGetValue = pRow["P30"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vConvertString);

                //31
                vGetValue = pRow["P31"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //32
                vGetValue = pRow["P32"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 53, vConvertString);

                //33
                vGetValue = pRow["P33"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 57, vConvertString);

                //34
                vGetValue = pRow["P34"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 61, vConvertString);


                ///////////////////////////////////////////////////////////
                ///3번째///
                ///
                vXLine++;
                //35
                vGetValue = pRow["P35"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 1, vConvertString);
                }
                else
                {
                    mPrinting.XLCellMerge(vXLine, 1, vXLine + 2, 8, true);
                    mPrinting.XLSetCell(vXLine, 1, vConvertString);
                }

                //36
                vGetValue = pRow["P36"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                }

                //37번
                vGetValue = pRow["P37"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //38번
                vGetValue = pRow["P38"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //39번
                vGetValue = pRow["P39"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //40번
                vGetValue = pRow["P40"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //41번
                vGetValue = pRow["P41"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vConvertString);

                //42번
                vGetValue = pRow["P42"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //43번
                vGetValue = pRow["P43"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //44번
                vGetValue = pRow["P44"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

                //45
                vGetValue = pRow["P45"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vConvertString);

                //46
                vGetValue = pRow["P46"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vConvertString);

                //47
                vGetValue = pRow["P47"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vConvertString);

                //48
                vGetValue = pRow["P48"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //49
                vGetValue = pRow["P49"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 53, vConvertString);

                //50
                vGetValue = pRow["P50"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 57, vConvertString);

                //51
                vGetValue = pRow["P51"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 61, vConvertString);


                ///////////////////////////////////////////////////////////
                ///4번째///
                ///
                vXLine++;  
                //52
                vGetValue = pRow["P52"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                { 
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 1, vConvertString);
                } 

                //53
                vGetValue = pRow["P53"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                } 

                //54번
                vGetValue = pRow["P54"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //55
                vGetValue = pRow["P55"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //56
                vGetValue = pRow["P56"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //57
                vGetValue = pRow["P57"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //58
                vGetValue = pRow["P58"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vConvertString);

                //59
                vGetValue = pRow["P59"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //60
                vGetValue = pRow["P60"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //61
                vGetValue = pRow["P61"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

                //62
                vGetValue = pRow["P62"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vConvertString);

                //63
                vGetValue = pRow["P63"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vConvertString);

                //64
                vGetValue = pRow["P64"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vConvertString);

                //65
                vGetValue = pRow["P65"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //66
                vGetValue = pRow["P66"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 53, vConvertString);

                //67
                vGetValue = pRow["P67"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 57, vConvertString);

                //68
                vGetValue = pRow["P68"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 61, vConvertString);


                ///////////////////////////////////////////////////////////
                ///5번째///
                ///
                vXLine++;
                //69
                vGetValue = pRow["P69"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 1, vConvertString);
                } 

                //70
                vGetValue = pRow["P70"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                } 

                //71
                vGetValue = pRow["P71"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vConvertString);

                //72
                vGetValue = pRow["P72"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //73
                vGetValue = pRow["P73"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //74
                vGetValue = pRow["P74"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //75
                vGetValue = pRow["P75"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vConvertString);

                //76
                vGetValue = pRow["P76"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //77
                vGetValue = pRow["P77"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //78
                vGetValue = pRow["P78"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

                //79
                vGetValue = pRow["P79"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 37, vConvertString);

                //80
                vGetValue = pRow["P80"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 41, vConvertString);

                //81
                vGetValue = pRow["P81"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 45, vConvertString);

                //82
                vGetValue = pRow["P82"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //83
                vGetValue = pRow["P83"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 53, vConvertString);

                //84
                vGetValue = pRow["P84"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 57, vConvertString);

                //85
                vGetValue = pRow["P85"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert != true)
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 61, vConvertString);

                //총합계 및 부서 합계 배경색 변경.
                if (vSUMMARY_FLAG == "N")
                {
                    /////////
                }
                else
                {
                    //2.BACK COLOR 변경.
                    mPrinting.XLCellColorBrush(mCurrentRow, 2, mCurrentRow + 4, mCopy_EndCol - 1, System.Drawing.Color.LightBlue);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            }
            return vXLine;
        }

        private int XlLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pPrintingLine, int[] pGridColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vGetValue = null;
            object vGetValue_2 = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            decimal vConvertDecimal2 = 0m;
            
            decimal vTEMP_AMT = 0m;
            string vSUMMARY_FLAG = "N";

            bool IsConvert = false;
            bool IsConvert2 = false;
            try
            {
                vSUMMARY_FLAG = iConv.ISNull(pGrid.GetCellValue(pRow, pGridColumn[80]));

                mPrinting.XLActiveSheet(mTargetSheet);

                //[부서] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[0]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {
                    
                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {

                }
                else
                {
                    mPrinting.XLCellMerge(vXLine, 2, vXLine + 1, 7, true);                    
                }
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //[직위] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[1]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 1, 2, vConvertString);
                }

                //[입사일자] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[5]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 2, 2, vConvertString);
                }

                //[사번] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[2]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty; 
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                }

                
                //[성명] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[3]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 1, 5, vConvertString);
                }
                else
                {
                    mPrinting.XLCellMerge(vXLine + 2, 2, vXLine + 4, 7, true);
                    mPrinting.XLSetCell(vXLine + 2, 2, vConvertString);
                }

                //[급여구분] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[4]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 2, 5, vConvertString);
                }

                //[퇴사일자] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[6]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 2, vConvertString);

                //근무내역 및 기본사항.
                //[월정급]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[9]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vConvertDecimal);                    
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 8, vConvertString);

                //[연장시간]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[15]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 8, vConvertString);

                //[유휴]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[77]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 8, vConvertString);

                //[정상근무]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[11]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vConvertString);

                //[야간시간]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[21]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 11, vConvertString);

                //[무휴]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[78]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 11, vConvertString);

                //[근무(공가)]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[73]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //[휴일근로-토]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[17]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 14, vConvertString);

                //[근태공제]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[13]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 14, vConvertString);

                //[미근무]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[74]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vConvertString);

                //[휴일근로-일]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[19]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 17, vConvertString);

                //[주차]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[75]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vConvertString);

                //[미주차]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[76]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 20, vConvertString);

                //[통상시급]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[10]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 20, vConvertString);
                 
                ////////////////////////////////////////////////////////////////////////////////지급항목 
                //[기본급]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[12]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vConvertString);

                //휴일근로+휴일연장 수당 / 관리직 연장수당
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[18]);
                vGetValue_2 = pGrid.GetCellValue(pRow, pGridColumn[20]);
                 
                vTEMP_AMT = iConv.ISDecimaltoZero(vGetValue) + iConv.ISDecimaltoZero(vGetValue_2);
                IsConvert = IsConvertNumber(vTEMP_AMT, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vTEMP_AMT);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 23, vConvertString);

                //[기타수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[26]);
                vGetValue_2 = pGrid.GetCellValue(pRow, pGridColumn[28]);
                vTEMP_AMT = iConv.ISDecimaltoZero(vGetValue) + iConv.ISDecimaltoZero(vGetValue_2);
                IsConvert = IsConvertNumber(vTEMP_AMT, out vConvertDecimal); 
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vTEMP_AMT);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 23, vConvertString);

                //[상여금]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[32]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 23, vConvertString);

                ////////////27
                //[차량보조]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[31]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 27, vConvertString);

                //[당직비]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[29]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine+1, 27, vConvertString);

                //[연차수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[30]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 27, vConvertString);

                //[기타수당(포상금)]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[35]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 27, vConvertString);

                ////// 31
                //[연장수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[16]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine , 31, vConvertString);

                
                //[기타수당2]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[34]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine+1, 31, vConvertString);

                //[근태공제]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[14]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 31, vConvertString);
                
                
                //[심야수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[22]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vConvertString);

                //[특별수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[24]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine+1, 35, vConvertString);

                //[급여소급분]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[33]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 35, vConvertString);

                //[그외수당]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[27]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 4, 35, vConvertString);

                ///////////////////////////////////////////////////////////////////공제항목//

                //[국민연금]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[43]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 39, vConvertString);

                //[건강보험퇴직]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[59]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 39, vConvertString);

                //[건강검진]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[53]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 39, vConvertString);

                //[연말정산지방소득세]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[66]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 39, vConvertString);

                //[연말정산지방소득세]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[65]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 4, 39, vConvertString);

                //[건강보험]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[45]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 43, vConvertString);

                //[건강보험연말]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[51]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 43, vConvertString);

                //[장기요양보험]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[61]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 43, vConvertString);

                //[기타공제]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[46]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 43, vConvertString);

                //[가압류]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[62]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 4, 43, vConvertString);

                //[고용보험]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[44]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 47, vConvertString);


                //[식대공제]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[68]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 47, vConvertString);

                //[장기요양보험정산]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[48]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 47, vConvertString);


                //[중도정산]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[56]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 47, vConvertString);

                //[중도정산]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[54]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 4, 47, vConvertString);


                //[국민연금소급분]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[63]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 51, vConvertString);

                //[경조금]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[58]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 51, vConvertString);


                //[소득세]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[41]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 51, vConvertString);

                //[학자금공제]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[55]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 51, vConvertString);
                ////[건강연말정산]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[47]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 51, vConvertString);

                ////[연말정산농특세]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[67]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 51, vConvertString);

                //[건강보험정산]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[47]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 55, vConvertString);

                //[작업복]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[52]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 55, vConvertString);
                                
                //[지방소득세]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[42]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 55, vConvertString);

                //[전월정산액]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[50]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 55, vConvertString);
                //합계
                //[총지급액]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[40]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 59, vConvertString);

                //[공제총액]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[70]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 59, vConvertString);

                //[실지급액]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[71]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 59, vConvertString);

                ////[]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[67]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 55, vConvertString);

                //총합계 및 부서 합계 배경색 변경.
                if (vSUMMARY_FLAG == "N")
                {
                    /////////
                }
                else
                {
                    //2.BACK COLOR 변경.
                    mPrinting.XLCellColorBrush(mCurrentRow, 2, mCurrentRow + 4, mCopy_EndCol - 1, System.Drawing.Color.LightBlue);
                }

                vXLine = vXLine + 5;
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

         
        public int XLWirteMain(InfoSummit.Win.ControlAdv.ISDataAdapter pPrompt, InfoSummit.Win.ControlAdv.ISDataAdapter pApt,
                                object pLocal_DATE, object pUserName, object pCorporationName, object pYYYYMM, object pWageTypeName, object pDepartmentName)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
              
            //초기화//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 65;
            mCopy_EndRow = 48;

            //mDefaultEndPageRow = 1;
            mDefaultPageRow = 13;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 43;  //최종 인쇄 라인.
            //m1stPrintingLastRow = 40;

            mCurrentRow = 13;
            mCopyLineSUM = 1;

            int vCurrRow = 0;
            int vTotalRow = 0;
            int vPageRowCount = 0;  //인쇄후 해당 라인 증가 위해. 

            mCorporationName = pCorporationName;
            mUserName = pUserName;
            mYYYYMM = pYYYYMM;
            mWageTypeName = pWageTypeName;
            mDepartmentName = pDepartmentName;
            mPringingDateTime = pLocal_DATE;

            string vDEPT_CODE = string.Empty;
            object vDEPT_NAME = string.Empty;

            try
            {
                if (pPrompt.CurrentRows.Count > 0)
                    XlPrompt(pPrompt);

                vTotalRow = pApt.CurrentRows.Count;
                //TotalPage(pGrid);

                if (vTotalRow > 0)
                { 
                    vPageRowCount = mCurrentRow - 5;

                    foreach(System.Data.DataRow vRow in pApt.CurrentRows)
                    {
                        vCurrRow++;
                        vMessage = string.Format("Row : {0} / {1}", vCurrRow, vTotalRow);
                        mAppInterface.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        if (iConv.ISNull(vRow["SUMMARY_FLAG"]) == "T")
                        {
                            vDEPT_NAME = string.Empty;
                        }
                        else
                        {
                            vDEPT_NAME = iConv.ISNull(vRow["DEPT_NAME"]);
                        }
                        if (vCurrRow == 1)
                        {
                            //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pGrid, vRow, vDEPT_NAME);
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME);
                        }
                        else if (vDEPT_CODE != iConv.ISNull(vRow["DEPT_CODE"]) && mIsNewPage == false)
                        {
                            XlAllLineClear(pCorporationName);
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME);
                            //아직인쇄 전 이므로 페이지ROW에 +4를 해줌.
                            mCurrentRow = mCurrentRow + (mCopy_EndRow - (vPageRowCount + 5)) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                            vPageRowCount = mDefaultPageRow - 5;
                        }

                        mCurrentRow = XlLine(vRow, mCurrentRow);
                        vPageRowCount = vPageRowCount + 5;
                        if (iConv.ISNull(vRow["SUMMARY_FLAG"]) == "T")
                        {

                        }
                        else
                        {
                            vDEPT_CODE = iConv.ISNull(vRow["DEPT_CODE"]);
                        }

                        if (vCurrRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            SumWrite(mCurrentRow);      //합계.
                            if (vPageRowCount != mPrintingLastRow)
                            {
                                //마지막ROW가 마지막 인쇄하고 다르면 엑셀 라인 CLEAR
                                XlAllLineClear(pCorporationName);
                            }
                        }
                        else
                        {
                            IsNewPage(vPageRowCount, false, vDEPT_NAME);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                //인쇄 후 이므로 현재 페이지ROW에 -4를 해줌.
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 5) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow - 5;
                            }
                        }
                    } 
                    mPrinting.XLDeleteSheet(mSourceSheet1); 
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
            int vPageCount = 48;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageCount * (r - 1);
                mPrinting.XLSetCell((vLINE + 4), 58, string.Format("Page {0} of {1}", r, mPageNumber)); 
            } 
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, object pDEPT_NAME)
        {
            if (mPrintingLastRow == pPrintingLine)
            {
                mIsNewPage = true;                
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true; 
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
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

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, object pDEPT_NAME)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);
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
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopySumPrintingLine));
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, object pDEPT_NAME)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);            
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
            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mCorporationName);
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
            mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

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
            mPrinting.XLPrinting(pPageSTART, pPageEND, 1);
        }

        public void PreviewPrinting(int pPageSTART, int pPageEND)
        { 
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