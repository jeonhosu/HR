using System;

namespace HRMF0537
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterfaceAdv = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mCopySumPrintingLine = 1; //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치
        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;
        private int mPrintingLineMAX = 43; //43:12행부터 43행까지, 32:반복되는 값이
        private int mIncrementCopyMAX = 45;
        private int mPositionPrintLineSTART = 12; //라인 출력시 엑셀 시작 행 위치 지정

        private string mCorporationName = string.Empty;
        private string mUserName = string.Empty;
        private string mYYYYMM = string.Empty;
        private string mWageTypeName = string.Empty;
        private string mDepartmentName = string.Empty;
        private string mPringingDateTime = string.Empty;

        private string mPageString = string.Empty;
        private int mPageTotalNumber = 0;
        private int mCountPage = 0;

        private string[] mGridColumn;
        private int[] mXLColumn;
        private decimal[] mSumValueColumn;

        private string mDepartment = string.Empty;

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

        public int PrintingLineMAX
        {
            set
            {
                mPrintingLineMAX = value;
            }
        }

        public int IncrementCopyMAX
        {
            set
            {
                mIncrementCopyMAX = value;
            }
        }

        public int PositionPrintLineSTART
        {
            set
            {
                mPositionPrintLineSTART = value;
            }
        }

        public int CopySumPrintingLine
        {
            set
            {
                mCopySumPrintingLine = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterfaceAdv)
        {
            mPrinting = new XL.XLPrint();
            mAppInterfaceAdv = pAppInterfaceAdv;
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

        private void XlAllLineClear(int[] pXLColumn)
        {
            object vObject = null;
            int vPrintingLineMAX = mPrintingLineMAX + 1;
            int mXLColumnCount = pXLColumn.Length;

            mPrinting.XLActiveSheet("SourceTab1");

            for (int vXLine = mPositionPrintLineSTART; vXLine < vPrintingLineMAX; vXLine++)
            {
                for (int vCOL = 1; vCOL < mXLColumnCount; vCOL++)
                {
                    mPrinting.XLSetCell(vXLine, pXLColumn[vCOL], vObject);
                }
            }
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(string pUserName, string pPrintingDateTime, string pYYYYMM, string pWageTypeName, string pDepartment_NAME, string pPageString, string pCorporationName)
        {
            bool isNull = false;
            try
            {
                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(2, 2);    //Title
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(4, 6);    //출력자
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(5, 6);    //급여구분
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 19);   //부서
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(4, 56);   //페이지
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(5, 56);   //출력일자
                System.Drawing.Point vCellPoint07 = new System.Drawing.Point(44, 41);  //업체

                mPrinting.XLActiveSheet("SourceTab1"); //셀에 문자를 넣기 위해 쉬트 선택

                //Title
                isNull = string.IsNullOrEmpty(pYYYYMM);
                if (isNull != true)
                {
                    string vYear = pYYYYMM.Substring(0, 4);
                    string vMonth = pYYYYMM.Substring(5, 2);
                    string vTitle = string.Format("{0}년 {1}월 급여 대장", vYear, vMonth);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vTitle);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, null);
                }

                //출력자
                isNull = string.IsNullOrEmpty(pUserName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, pUserName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, null);
                }

                //급여구분
                isNull = string.IsNullOrEmpty(pWageTypeName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, pWageTypeName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, "전체");
                }

                //부서
                isNull = string.IsNullOrEmpty(pDepartment_NAME);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, pDepartment_NAME);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, "전체");
                }

                //페이지
                isNull = string.IsNullOrEmpty(pPageString);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, pPageString);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, null);
                }

                //출력일자
                isNull = string.IsNullOrEmpty(pPrintingDateTime);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, pPrintingDateTime);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, null);
                }

                //업체
                isNull = string.IsNullOrEmpty(pCorporationName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, pCorporationName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, null);
                }
            }
            catch (System.Exception ex)
            {
                mAppInterfaceAdv.OnAppMessage(ex.Message);

                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pGridColumn, out int[] pXLColumn)
        {
            mSumValueColumn = new decimal[69];
            pGridColumn = new string[69];
            pXLColumn = new int[69];

            pGridColumn[01] = "DEPT_NAME";                    //부서
            pGridColumn[02] = "PERSON_NUM";                   //사원번호
            pGridColumn[03] = ""; //기본급
            pGridColumn[04] = "TOTAL_ATT_DAY"; //정산근무
            pGridColumn[05] = "HOLY_1_TIME"; //휴일특근
            pGridColumn[06] = "DUTY_30"; //근무(공가)
            pGridColumn[07] = "TOT_DED_COUNT"; //미근무
            pGridColumn[08] = "A01";                          //기본급[지급항목]
            pGridColumn[09] = "A08";                          //가족수당
            pGridColumn[10] = "A02";                          //직책수당
            pGridColumn[11] = "A03";                          //근속수당
            pGridColumn[12] = "A18"; //연차수당
            pGridColumn[13] = "D01";                          //소득세
            pGridColumn[14] = "D02";                          //주민세
            pGridColumn[15] = "D03";                          //국민연금
            pGridColumn[16] = "D05";                          //건강보험
            pGridColumn[17] = "TOT_SUPPLY_AMOUNT";            //총지급액[총지급액]
            pGridColumn[18] = "POST_NAME";                    //직위
            pGridColumn[19] = "NAME";                         //성명
            pGridColumn[20] = ""; //일급
            pGridColumn[21] = "OVER_TIME"; //연장근로(연장시간)
            pGridColumn[22] = "HOLY_1_OT"; //휴일연장
            pGridColumn[23] = "S_HOLY_1_COUNT"; //주차
            pGridColumn[24] = "WEEKLY_DED_COUNT"; //미주차
            pGridColumn[25] = "A06";                          //자격수당
            pGridColumn[26] = "A11";                          //시간외수당
            pGridColumn[27] = "A12";                          //연장수당
            pGridColumn[28] = "A13";                          //야간수당
            pGridColumn[29] = "A14";                          //특근수당
            pGridColumn[30] = "D04";                          //고용보험
            pGridColumn[31] = "";                          //사우회
            pGridColumn[32] = "D06"; //식대
            pGridColumn[33] = ""; //가불금
            pGridColumn[34] = "TOT_DED_AMOUNT";               //총공제액[총공제액]
            pGridColumn[35] = ""; //정식입사일
            pGridColumn[36] = "WAGE_TYPE_NAME";               //급상여구분
            pGridColumn[37] = ""; //시급
            pGridColumn[38] = "NIGHT_BONUS_TIME"; //야간근로(야간시간)
            pGridColumn[39] = "HOLY_1_NIGHT"; //휴일야간
            pGridColumn[40] = "HOLY_1_COUNT"; //유휴
            pGridColumn[41] = "DUTY_30"; //공가
            pGridColumn[42] = "";                          //만근수당
            pGridColumn[43] = "";                          //지각외출조퇴
            pGridColumn[44] = "";                          //결근
            pGridColumn[45] = "A25";                          //차량유지비
            pGridColumn[46] = "A07";                          //기타수당
            pGridColumn[47] = ""; //경조사비
            pGridColumn[48] = ""; //카드대
            pGridColumn[49] = "D14";                          //기타
            pGridColumn[50] = ""; //동호회
            pGridColumn[51] = "REAL_AMOUNT";                  //실지급액[실지급액]
            pGridColumn[52] = ""; //
            pGridColumn[53] = ""; //
            pGridColumn[54] = ""; //
            pGridColumn[55] = ""; //
            pGridColumn[56] = "LATE_TIME"; //근태공제
            pGridColumn[57] = "HOLY_0_COUNT"; //무휴
            pGridColumn[58] = ""; //통상시급
            pGridColumn[59] = ""; //
            pGridColumn[60] = "A09";                          //상여금
            pGridColumn[61] = "";                          //야간장려수당
            pGridColumn[62] = "A10";                          //전월소급분
            pGridColumn[63] = ""; //무급휴가
            pGridColumn[64] = "D16";                          //정산소득세
            pGridColumn[65] = "D17";                          //정산주민세
            pGridColumn[66] = ""; //국민연금소급분
            pGridColumn[67] = ""; //건강연말정산
            pGridColumn[68] = ""; //비과세총액

            pXLColumn[01] = 2;  //부서
            pXLColumn[02] = 5;  //사원번호
            pXLColumn[03] = 8;  //기본급
            pXLColumn[04] = 11; //정산근무
            pXLColumn[05] = 14; //휴일특근
            pXLColumn[06] = 17; //근무(공가)
            pXLColumn[07] = 20; //미근무
            pXLColumn[08] = 23; //기본급[지급항목]
            pXLColumn[09] = 27; //가족수당
            pXLColumn[10] = 31; //직책수당
            pXLColumn[11] = 35; //근속수당
            pXLColumn[12] = 39; //연차수당
            pXLColumn[13] = 43; //소득세
            pXLColumn[14] = 47; //주민세
            pXLColumn[15] = 51; //국민연금
            pXLColumn[16] = 55; //건강보험
            pXLColumn[17] = 59; //총지급액[총지급액]
            pXLColumn[18] = 2;  //직위
            pXLColumn[19] = 5;  //성명
            pXLColumn[20] = 8;  //일급
            pXLColumn[21] = 11; //연장근로
            pXLColumn[22] = 14; //휴일연장
            pXLColumn[23] = 17; //주차
            pXLColumn[24] = 20; //미주차
            pXLColumn[25] = 23; //자격수당
            pXLColumn[26] = 27; //시간외수당
            pXLColumn[27] = 31; //년차수당
            pXLColumn[28] = 35; //야간수당
            pXLColumn[29] = 39; //특근수당
            pXLColumn[30] = 43; //고용보험
            pXLColumn[31] = 47; //사우회
            pXLColumn[32] = 51; //식대
            pXLColumn[33] = 55; //가불금
            pXLColumn[34] = 59; //총공제액[총공제액]
            pXLColumn[35] = 2;  //정식입사일
            pXLColumn[36] = 5;  //급상여구분
            pXLColumn[37] = 8;  //시급
            pXLColumn[38] = 11; //야간근로
            pXLColumn[39] = 14; //휴일야간
            pXLColumn[40] = 17; //유휴
            pXLColumn[41] = 20; //공가
            pXLColumn[42] = 23; //만근수당
            pXLColumn[43] = 27; //지각외출조퇴
            pXLColumn[44] = 31; //결근
            pXLColumn[45] = 35; //차량유지비
            pXLColumn[46] = 39; //기타수당
            pXLColumn[47] = 43; //경조사비
            pXLColumn[48] = 47; //카드대
            pXLColumn[49] = 51; //기타
            pXLColumn[50] = 55; //동호회
            pXLColumn[51] = 59; //실지급액[실지급액]
            pXLColumn[52] = 2;  //
            pXLColumn[53] = 5;  //
            pXLColumn[54] = 8;  //
            pXLColumn[55] = 11; //
            pXLColumn[56] = 14; //근태공제
            pXLColumn[57] = 17; //무휴
            pXLColumn[58] = 20; //통상시급
            pXLColumn[59] = 23; //
            pXLColumn[60] = 27; //상여금
            pXLColumn[61] = 31; //야간장려수당
            pXLColumn[62] = 35; //전월소급분
            pXLColumn[63] = 39; //무급휴가
            pXLColumn[64] = 43; //정산소득세
            pXLColumn[65] = 47; //정산주민세
            pXLColumn[66] = 51; //국민연금소급분
            pXLColumn[67] = 55; //건강연말정산
            pXLColumn[68] = 59; //비과세총액
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
                mAppInterfaceAdv.OnAppMessage(mMessageError);
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
                mAppInterfaceAdv.OnAppMessage(mMessageError);
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
                mAppInterfaceAdv.OnAppMessage(mMessageError);
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

        #region ----- Xl Clear Value SUM Methods -----

        private void ClearValueSumValue()
        {
            int vCountIndex = mSumValueColumn.Length;

            for (int vRow = 0; vRow < vCountIndex; vRow++)
            {
                mSumValueColumn[vRow] = 0m;
            }
        }

        #endregion;

        #region ----- Xl SUM Methods -----

        private int XLSUM(int pPrintingLine, int[] pXLColumn, string pDepartment, decimal[] pSumValueColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호
            int vXLIndexColumn = 0;
            string vConvertString = string.Empty;
            decimal vSumValue = 0m;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                //[01]
                vXLIndexColumn = pXLColumn[1];
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, pDepartment);

                //[03]
                vXLIndexColumn = pXLColumn[3];
                vSumValue = pSumValueColumn[3];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[04]
                vXLIndexColumn = pXLColumn[4];
                vSumValue = pSumValueColumn[4];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[05]
                vXLIndexColumn = pXLColumn[5];
                vSumValue = pSumValueColumn[5];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[06]
                vXLIndexColumn = pXLColumn[6];
                vSumValue = pSumValueColumn[6];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[07]
                vXLIndexColumn = pXLColumn[7];
                vSumValue = pSumValueColumn[7];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[08]
                vXLIndexColumn = pXLColumn[8];
                vSumValue = pSumValueColumn[8];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[09]
                vXLIndexColumn = pXLColumn[9];
                vSumValue = pSumValueColumn[9];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[10]
                vXLIndexColumn = pXLColumn[10];
                vSumValue = pSumValueColumn[10];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[11]
                vXLIndexColumn = pXLColumn[11];
                vSumValue = pSumValueColumn[11];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[12]
                vXLIndexColumn = pXLColumn[12];
                vSumValue = pSumValueColumn[12];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[13]
                vXLIndexColumn = pXLColumn[13];
                vSumValue = pSumValueColumn[13];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[14]
                vXLIndexColumn = pXLColumn[14];
                vSumValue = pSumValueColumn[14];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[15]
                vXLIndexColumn = pXLColumn[15];
                vSumValue = pSumValueColumn[15];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[16]
                vXLIndexColumn = pXLColumn[16];
                vSumValue = pSumValueColumn[16];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[17]
                vXLIndexColumn = pXLColumn[17];
                vSumValue = pSumValueColumn[17];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[20]
                vXLIndexColumn = pXLColumn[20];
                vSumValue = pSumValueColumn[20];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[21]
                vXLIndexColumn = pXLColumn[21];
                vSumValue = pSumValueColumn[21];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[22]
                vXLIndexColumn = pXLColumn[22];
                vSumValue = pSumValueColumn[22];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[23]
                vXLIndexColumn = pXLColumn[23];
                vSumValue = pSumValueColumn[23];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[24]
                vXLIndexColumn = pXLColumn[24];
                vSumValue = pSumValueColumn[24];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[25]
                vXLIndexColumn = pXLColumn[25];
                vSumValue = pSumValueColumn[25];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[26]
                vXLIndexColumn = pXLColumn[26];
                vSumValue = pSumValueColumn[26];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[27]
                vXLIndexColumn = pXLColumn[27];
                vSumValue = pSumValueColumn[27];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[28]
                vXLIndexColumn = pXLColumn[28];
                vSumValue = pSumValueColumn[28];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[29]
                vXLIndexColumn = pXLColumn[29];
                vSumValue = pSumValueColumn[29];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[30]
                vXLIndexColumn = pXLColumn[30];
                vSumValue = pSumValueColumn[30];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[31]
                vXLIndexColumn = pXLColumn[31];
                vSumValue = pSumValueColumn[31];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[32]
                vXLIndexColumn = pXLColumn[32];
                vSumValue = pSumValueColumn[32];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[33]
                vXLIndexColumn = pXLColumn[33];
                vSumValue = pSumValueColumn[33];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[34]
                vXLIndexColumn = pXLColumn[34];
                vSumValue = pSumValueColumn[34];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[37]
                vXLIndexColumn = pXLColumn[37];
                vSumValue = pSumValueColumn[37];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[38]
                vXLIndexColumn = pXLColumn[38];
                vSumValue = pSumValueColumn[38];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[39]
                vXLIndexColumn = pXLColumn[39];
                vSumValue = pSumValueColumn[39];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[40]
                vXLIndexColumn = pXLColumn[40];
                vSumValue = pSumValueColumn[40];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[41]
                vXLIndexColumn = pXLColumn[41];
                vSumValue = pSumValueColumn[41];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[42]
                vXLIndexColumn = pXLColumn[42];
                vSumValue = pSumValueColumn[42];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[43]
                vXLIndexColumn = pXLColumn[43];
                vSumValue = pSumValueColumn[43];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[44]
                vXLIndexColumn = pXLColumn[44];
                vSumValue = pSumValueColumn[44];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[45]
                vXLIndexColumn = pXLColumn[45];
                vSumValue = pSumValueColumn[45];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[46]
                vXLIndexColumn = pXLColumn[46];
                vSumValue = pSumValueColumn[46];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[47]
                vXLIndexColumn = pXLColumn[47];
                vSumValue = pSumValueColumn[47];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[48]
                vXLIndexColumn = pXLColumn[48];
                vSumValue = pSumValueColumn[48];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[49]
                vXLIndexColumn = pXLColumn[49];
                vSumValue = pSumValueColumn[49];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[50]
                vXLIndexColumn = pXLColumn[50];
                vSumValue = pSumValueColumn[50];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[51]
                vXLIndexColumn = pXLColumn[51];
                vSumValue = pSumValueColumn[51];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[54]
                vXLIndexColumn = pXLColumn[54];
                vSumValue = pSumValueColumn[54];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[55]
                vXLIndexColumn = pXLColumn[55];
                vSumValue = pSumValueColumn[55];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[56]
                vXLIndexColumn = pXLColumn[56];
                vSumValue = pSumValueColumn[56];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[57]
                vXLIndexColumn = pXLColumn[57];
                vSumValue = pSumValueColumn[57];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[58]
                vXLIndexColumn = pXLColumn[58];
                vSumValue = pSumValueColumn[58];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[59]
                vXLIndexColumn = pXLColumn[59];
                vSumValue = pSumValueColumn[59];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[60]
                vXLIndexColumn = pXLColumn[60];
                vSumValue = pSumValueColumn[60];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[61]
                vXLIndexColumn = pXLColumn[61];
                vSumValue = pSumValueColumn[61];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[62]
                vXLIndexColumn = pXLColumn[62];
                vSumValue = pSumValueColumn[62];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[63]
                vXLIndexColumn = pXLColumn[63];
                vSumValue = pSumValueColumn[63];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[64]
                vXLIndexColumn = pXLColumn[64];
                vSumValue = pSumValueColumn[64];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[65]
                vXLIndexColumn = pXLColumn[65];
                vSumValue = pSumValueColumn[65];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[66]
                vXLIndexColumn = pXLColumn[66];
                vSumValue = pSumValueColumn[66];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[67]
                vXLIndexColumn = pXLColumn[67];
                vSumValue = pSumValueColumn[67];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[68]
                vXLIndexColumn = pXLColumn[68];
                vSumValue = pSumValueColumn[68];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }


            pPrintingLine = vXLine;
            IsNewPage(pPrintingLine);
            if (mIsNewPage == true)
            {
                pPrintingLine = mPositionPrintLineSTART;
            }

            return pPrintingLine;
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pPrintingLine, string[] pGridColumn, int[] pXLColumn)
        {
            bool vIsValueViewTemp = true;
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vGetValue = null;
            int vGridIndexColumn = 0;
            int vXLIndexColumn = 0;

            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                //[01]
                vXLIndexColumn = pXLColumn[1];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[1]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[01]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[01]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[02]
                vXLIndexColumn = pXLColumn[2];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[2]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[02]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[02]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[03]
                vXLIndexColumn = pXLColumn[3];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[3]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[03]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[3] = mSumValueColumn[3] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[03]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[03]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[04]
                vXLIndexColumn = pXLColumn[4];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[4]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[04]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[4] = mSumValueColumn[4] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[04]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[04]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[05]
                vXLIndexColumn = pXLColumn[5];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[5]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[05]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[5] = mSumValueColumn[5] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[05]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[05]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[06]
                vXLIndexColumn = pXLColumn[6];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[6]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[06]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[6] = mSumValueColumn[6] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[06]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[06]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[07]
                vXLIndexColumn = pXLColumn[7];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[7]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[07]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[7] = mSumValueColumn[7] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[07]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[07]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[08]
                vXLIndexColumn = pXLColumn[8];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[8]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[8] = mSumValueColumn[8] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[08]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[08]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[09]
                vXLIndexColumn = pXLColumn[9];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[9]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[9] = mSumValueColumn[9] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[09]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[09]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[10]
                vXLIndexColumn = pXLColumn[10];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[10]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[10] = mSumValueColumn[10] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[10]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[10]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[11]
                vXLIndexColumn = pXLColumn[11];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[11]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[11] = mSumValueColumn[11] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[11]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[11]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[12]
                vXLIndexColumn = pXLColumn[12];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[12]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[12] = mSumValueColumn[12] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[12]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[12]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[13]
                vXLIndexColumn = pXLColumn[13];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[13]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[13] = mSumValueColumn[13] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[13]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[13]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[14]
                vXLIndexColumn = pXLColumn[14];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[14]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[14] = mSumValueColumn[14] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[14]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[14]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[15]
                vXLIndexColumn = pXLColumn[15];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[15]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[15] = mSumValueColumn[15] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[15]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[15]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[16]
                vXLIndexColumn = pXLColumn[16];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[16]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[16] = mSumValueColumn[16] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[16]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[16]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[17]
                vXLIndexColumn = pXLColumn[17];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[17]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[17] = mSumValueColumn[17] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[17]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[17]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[18]
                vXLIndexColumn = pXLColumn[18];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[18]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[18]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[18]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[19]
                vXLIndexColumn = pXLColumn[19];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[19]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[19]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[19]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[20]
                vXLIndexColumn = pXLColumn[20];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[20]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[20]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[20] = mSumValueColumn[20] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[20]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[20]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[21]
                vXLIndexColumn = pXLColumn[21];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[21]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[21]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[21] = mSumValueColumn[21] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[21]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[21]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[22]
                vXLIndexColumn = pXLColumn[22];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[22]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[22]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[22] = mSumValueColumn[22] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[22]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[22]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[23]
                vXLIndexColumn = pXLColumn[23];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[23]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[23]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[23] = mSumValueColumn[23] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[23]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[23]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[24]
                vXLIndexColumn = pXLColumn[24];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[24]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[24]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[24] = mSumValueColumn[24] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[24]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[24]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[25]
                vXLIndexColumn = pXLColumn[25];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[25]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[25] = mSumValueColumn[25] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[25]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[25]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[26]
                vXLIndexColumn = pXLColumn[26];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[26]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[26] = mSumValueColumn[26] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[26]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[26]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[27]
                vXLIndexColumn = pXLColumn[27];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[27]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[27] = mSumValueColumn[27] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[27]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[27]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[28]
                vXLIndexColumn = pXLColumn[28];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[28]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[28] = mSumValueColumn[28] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[28]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[28]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[29]
                vXLIndexColumn = pXLColumn[29];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[29]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[29] = mSumValueColumn[29] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[29]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[29]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[30]
                vXLIndexColumn = pXLColumn[30];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[30]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[30] = mSumValueColumn[30] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[30]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[30]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[31]
                vXLIndexColumn = pXLColumn[31];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[31]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[31] = mSumValueColumn[31] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[31]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[31]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[32]
                vXLIndexColumn = pXLColumn[32];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[32]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[32] = mSumValueColumn[32] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[32]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[32]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[33]
                vXLIndexColumn = pXLColumn[33];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[33]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[33] = mSumValueColumn[33] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[33]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[33]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[34]
                vXLIndexColumn = pXLColumn[34];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[34]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[34] = mSumValueColumn[34] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[34]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[34]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[35]
                vXLIndexColumn = pXLColumn[35];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[35]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertDate(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[35]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[35]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[36]
                vXLIndexColumn = pXLColumn[36];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[36]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[36]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[36]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[37]
                vXLIndexColumn = pXLColumn[37];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[37]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[37]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[37] = mSumValueColumn[37] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[37]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[37]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[38]
                vXLIndexColumn = pXLColumn[38];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[38]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[38]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[38] = mSumValueColumn[38] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[38]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[38]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[39]
                vXLIndexColumn = pXLColumn[39];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[39]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[39]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[39] = mSumValueColumn[39] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[39]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[39]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[40]
                vXLIndexColumn = pXLColumn[40];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[40]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[40]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[40] = mSumValueColumn[40] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[40]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[40]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[41]
                vXLIndexColumn = pXLColumn[41];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[41]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[41]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[41] = mSumValueColumn[41] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[41]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[41]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[42]
                vXLIndexColumn = pXLColumn[42];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[42]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[42] = mSumValueColumn[42] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[42]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[42]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[43]
                vXLIndexColumn = pXLColumn[43];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[43]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[43] = mSumValueColumn[43] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[43]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[43]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[44]
                vXLIndexColumn = pXLColumn[44];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[44]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[44] = mSumValueColumn[44] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[44]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[44]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[45]
                vXLIndexColumn = pXLColumn[45];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[45]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[45] = mSumValueColumn[45] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[45]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[45]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[46]
                vXLIndexColumn = pXLColumn[46];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[46]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[46] = mSumValueColumn[46] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[46]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[46]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[47]
                vXLIndexColumn = pXLColumn[47];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[47]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[47] = mSumValueColumn[47] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[47]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[47]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[48]
                vXLIndexColumn = pXLColumn[48];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[48]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[48] = mSumValueColumn[48] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[48]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[48]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[49]
                vXLIndexColumn = pXLColumn[49];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[49]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[49] = mSumValueColumn[49] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[49]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[49]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[50]
                vXLIndexColumn = pXLColumn[50];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[50]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[50] = mSumValueColumn[50] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[50]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[50]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[51]
                vXLIndexColumn = pXLColumn[51];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[51]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[51] = mSumValueColumn[51] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[51]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[51]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[52]
                vXLIndexColumn = pXLColumn[52];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[52]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[52]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[52]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[53]
                vXLIndexColumn = pXLColumn[53];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[53]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[53]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[53]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[54]
                vXLIndexColumn = pXLColumn[54];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[54]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[54]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[54] = mSumValueColumn[54] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[54]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[54]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[55]
                vXLIndexColumn = pXLColumn[55];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[55]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[55]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[55] = mSumValueColumn[55] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[55]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[55]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[56]
                vXLIndexColumn = pXLColumn[56];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[56]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[56]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[56] = mSumValueColumn[56] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[56]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[56]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[57]
                vXLIndexColumn = pXLColumn[57];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[57]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[57]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[57] = mSumValueColumn[57] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[57]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[57]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[58]
                vXLIndexColumn = pXLColumn[58];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[58]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[58]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[58] = mSumValueColumn[58] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[58]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[58]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[59]
                vXLIndexColumn = pXLColumn[59];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[59]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[59] = mSumValueColumn[59] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[59]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[59]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[60]
                vXLIndexColumn = pXLColumn[60];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[60]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[60] = mSumValueColumn[60] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[60]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[60]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[61]
                vXLIndexColumn = pXLColumn[61];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[61]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[61] = mSumValueColumn[61] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[61]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[61]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[62]
                vXLIndexColumn = pXLColumn[62];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[62]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[62] = mSumValueColumn[62] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[62]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[62]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[63]
                vXLIndexColumn = pXLColumn[63];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[63]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[63] = mSumValueColumn[63] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[63]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[63]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[64]
                vXLIndexColumn = pXLColumn[64];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[64]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[64] = mSumValueColumn[64] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[64]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[64]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[65]
                vXLIndexColumn = pXLColumn[65];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[65]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[65] = mSumValueColumn[65] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[65]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[65]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[66]
                vXLIndexColumn = pXLColumn[66];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[66]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[66] = mSumValueColumn[66] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[66]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[66]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[67]
                vXLIndexColumn = pXLColumn[67];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[67]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[67] = mSumValueColumn[67] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[67]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[67]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[68]
                vXLIndexColumn = pXLColumn[68];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[68]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[68] = mSumValueColumn[68] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[68]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[68]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }


            pPrintingLine = vXLine;
            IsNewPage(pPrintingLine);
            if (mIsNewPage == true)
            {
                pPrintingLine = mPositionPrintLineSTART;
            }

            return pPrintingLine;
        }

        #endregion;

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            if (mPrintingLineMAX < pPrintingLine)
            {
                mIsNewPage = true;
                mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);

                XlAllLineClear(mXLColumn);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory, string pUserName, string pCorporationName, string pYYYYMM, string pWageTypeName, string pDepartmentName)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            mPringingDateTime = string.Format("{0} {1}", vPrintingDate, vPrintingTime);

            int vPrintingLine = mPositionPrintLineSTART;

            object vObject = null;
            int vGridIndexColumn = 0;
            string vDepartment = string.Empty;

            try
            {
                int vTotalRow = pGrid.RowCount;
                mPageTotalNumber = vTotalRow / 8;
                mPageTotalNumber = (vTotalRow % 8) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);

                int vCountRow = 0;
                if (vTotalRow > 0)
                {
                    SetArray(out mGridColumn, out mXLColumn);

                    mCorporationName = pCorporationName;   //업체
                    mUserName = pUserName;                 //출력자
                    mYYYYMM = pYYYYMM;                     //출력년월
                    mWageTypeName = pWageTypeName;         //급여구분
                    mDepartmentName = pDepartmentName;     //부서

                    vObject = pGrid.GetCellValue(mGridColumn[1]);
                    mDepartment = ConvertString(vObject);

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vCountRow++;

                        vMessage = string.Format("Row : {0} / {1}", vRow, vTotalRow);
                        mAppInterfaceAdv.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vGridIndexColumn = pGrid.GetColumnToIndex(mGridColumn[1]);
                        vObject = pGrid.GetCellValue(vRow, vGridIndexColumn);
                        vDepartment = ConvertString(vObject);

                        //////[부서합계]
                        //////if (mDepartment != vDepartment)
                        //////{
                        //////    vPrintingLine = XLSUM(vPrintingLine, mXLColumn, mDepartment, mSumValueColumn);
                        //////    ClearValueSumValue();
                        //////    mDepartment = vDepartment;
                        //////}

                        vPrintingLine = XlLine(pGrid, vRow, vPrintingLine, mGridColumn, mXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            if (mPositionPrintLineSTART != vPrintingLine)
                            {
                                mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);
                            }

                            XlAllLineClear(mXLColumn);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mCountPage;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = 1; //복사되어질 쉬트의 폭, 시작열
            int vPrintHeaderColumnEND = 63;  //복사되어질 쉬트의 폭, 종료열

            mCountPage++;
            mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mPageString, mCorporationName);

            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            mPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mIncrementCopyMAX, vPrintHeaderColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}