using System;
using ISCommonUtil;

namespace HRMF0326
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SOURCE1"; 
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;  //최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 0;    // 페이지 증가후 PageCount 기본값.

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

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
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
        {// 파일명 뒤에 일련번호 증가.
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

        #region ----- Line SLIP Methods ----

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[6];

            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("SUPP_CUST_NAME");
            pGDColumn[1] = pGrid.GetColumnToIndex("TAX_REG_NO");
            pGDColumn[2] = pGrid.GetColumnToIndex("FORWARD_AMT");
            pGDColumn[3] = pGrid.GetColumnToIndex("INC_AMT");
            pGDColumn[4] = pGrid.GetColumnToIndex("DEC_AMT");
            pGDColumn[5] = pGrid.GetColumnToIndex("REMAIN_AMT");
        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[5];

            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("GL_DATE");
            pGDColumn[1] = pGrid.GetColumnToIndex("REMARKS");
            pGDColumn[2] = pGrid.GetColumnToIndex("DR_AMT");
            pGDColumn[3] = pGrid.GetColumnToIndex("CR_AMT");
            pGDColumn[4] = pGrid.GetColumnToIndex("REMAIN_AMT");
        }

        #endregion;

        #region ----- Array Set 2 ----

        //private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        //{// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
        //    pGDColumn = new int[3];
        //    pXLColumn = new int[3];
        //    // 그리드 or 아답터 위치.
        //    pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
        //    pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
        //    pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

        //    // 엑셀에 인쇄해야 할 위치.
        //    pXLColumn[0] = 20;
        //    pXLColumn[1] = 25;
        //    pXLColumn[2] = 30;
        //}

        #endregion;

        #region ----- Array Set 2  : Adapter 적용시 ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// 아답터의 table 값.
        //    pGDColumn = new int[10];
        //    pXLColumn = new int[10];

        //    pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
        //    pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
        //    pGDColumn[5] = pTable.Columns.IndexOf("PRICE_TERM_NAME");
        //    pGDColumn[6] = pTable.Columns.IndexOf("PAYMENT_METHOD_NAME");
        //    pGDColumn[7] = pTable.Columns.IndexOf("PAYMENT_TERM_NAME");
        //    pGDColumn[8] = pTable.Columns.IndexOf("REMARK");
        //    pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");


        //    pXLColumn[0] = 9;   //PO_TYPE_NAME
        //    pXLColumn[1] = 25;  //DISPLAY_NAME
        //    pXLColumn[2] = 42;  //PO_DATE
        //    pXLColumn[3] = 54;  //PO_NO
        //    pXLColumn[4] = 9;   //SUPPLIER_SHORT_NAME
        //    pXLColumn[5] = 35;  //PRICE_TERM_NAME
        //    pXLColumn[6] = 14;  //PAYMENT_METHOD_NAME
        //    pXLColumn[7] = 41;  //PAYMENT_TERM_NAME
        //    pXLColumn[8] = 7;   //REMARK
        //    pXLColumn[9] = 49;  //금액
        //}

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// 문자열 여부 체크 및 해당 값 리턴.
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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {// 숫자 여부 체크 및 해당 값 리턴.
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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {// 날짜 여부 체크 및 해당 값 리턴.
            bool vIsConvert = false;
            pConvertDateTimeShort = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            return vIsConvert;
        }

        #endregion;

        #region ----- Header Write Method ----

        public void HeaderWrite_1(object pCORP_NAME, string pPERIOD_NAME)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 기간
                vXLine = 3;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD_NAME);

                vXLine = 5;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, pCORP_NAME); 

                //인쇄일시//
                vXLine = 40;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintDateTime); 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_2(object pBUDGET_YEAR)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 기간
                vXLine = 3;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, pBUDGET_YEAR);

                vXLine = 33;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintDateTime);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 9; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                for (int i = 0; i <= pGrid.RowCount; i++)
                {
                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    vXLine++;
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Line Write Method -----

        private int XLLine_1(System.Data.DataRow pRow, int pXLine)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;
                        
            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m; 
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //근무일
                vObject = pRow["WORK_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                                
                //요일
                vObject = pRow["WORK_WEEK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString); 

                //작업장
                vObject = pRow["FLOOR_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-직위
                vObject = pRow["POST_NAME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-성명
                vObject = pRow["NAME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-근태
                vObject = pRow["DUTY_NAME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 22;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-근무
                vObject = pRow["HOLY_TYPE_NAME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-출근시간
                vObject = pRow["OPEN_TIME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2-퇴근시간
                vObject = pRow["CLOSE_TIME"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소정근로
                vObject = pRow["WORK_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 43;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지각조퇴
                vObject = pRow["LATE_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연장근로
                vObject = pRow["OVER_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 49;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //휴일근(토)
                vObject = pRow["HOLIDAY_0_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 52;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //휴일근로
                vObject = pRow["HOLIDAY_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //야간할증
                vObject = pRow["NIGHT_BONUS_TIME"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //비고
                vObject = pRow["DESCRIPTION"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 61;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }

        private int XLLine_2(System.Data.DataRow pRow, int pXLine, int pAccountRow, bool pPrint_Flag)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m; 
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //계정과목
                vObject = pRow["ACCOUNT_GROUP_DESC"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(pAccountRow, vXLColumn, vConvertString);

                if (pPrint_Flag == true)
                {
                    //계정과목
                    vObject = pRow["ACCOUNT_DESC"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    vXLColumn = 1;
                    mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                }
                
                //부서
                vObject = pRow["DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(pXLine, vXLColumn, vConvertString);
                
                //월별
                vObject = pRow["BUDGET_PERIOD"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //예산신청금액
                vObject = pRow["REQUEST_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 20;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배정금액
                vObject = pRow["ASSIGN_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //증감액
                vObject = pRow["VARY_AMOUNT"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //증가율
                vObject = pRow["VARY_RATE"];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal == 0)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###,###,###,###,###.00}", vConvertDecimal);
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //비고
                vObject = pRow["DESCRIPTION"];
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XL_TOTAL_Line(int pXLine, int[] pXLColumn)
        //{// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
        //    int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // 원본을 복사해서 타겟 에 복사해 넣음.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //10 - 보증금
        //        vXLColumnIndex = pXLColumn[10];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //11 - 월임대료
        //        vXLColumnIndex = pXLColumn[11];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //12 - 합계
        //        vXLColumnIndex = pXLColumn[12];
        //        IsConvert = IsConvertNumber(mTOT_LEASE_SUM_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //13 - 보증금이자
        //        vXLColumnIndex = pXLColumn[13];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_INTEREST_AMT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //14 - 월임대료(계)
        //        vXLColumnIndex = pXLColumn[14];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_SUM_AMT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //-------------------------------------------------------------------
        //        vXLine = vXLine + 1;
        //        //-------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mAppInterface.OnAppMessageEvent(mMessageError);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    pXLine = vXLine;

        //    return pXLine;
        //}

        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite_1(object pCORP_NAME, string pPERIOD_NAME, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty; 
            string vPERSON_NUM = string.Empty;

            int vRealPrinting = 10;     //30장마다 인쇄 처리.

            int vLIneRow = 0;
            int vTotalRow = 0;
            int vPageRowCount = 0; 

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 66;
            mCopy_EndRow = 40;
            mPrintingLastRow = 39;  //최종 인쇄 라인.

            mCurrentRow = 8;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 7;    // 페이지 증가후 PageCount 기본값.
            
            try
            {
                bool isOpen = XLFileOpen();

                //헤더 인쇄
                HeaderWrite_1(pCORP_NAME, pPERIOD_NAME);

                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pAdapter.OraSelectData.Rows.Count;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                #region ----- Header Write ----
                
                //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                //XLHeader1(pGrid, vGDColumn, vXLColumn);  // 헤더 인쇄.

                #endregion;

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                    {
                        vLIneRow++;

                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        //부서코드 동일여부 체크하여 다르면 page skip.
                        if (vPERSON_NUM == null || vPERSON_NUM == string.Empty)
                        {

                        }
                        else if (iString.ISNull(vRow["PERSON_NUM"], "-") == "-")
                        {

                        }
                        else if (vPERSON_NUM != iString.ISNull(vRow["PERSON_NUM"], "-"))
                        {
                            if (vRealPrinting < mPageNumber)
                            {
                                Printing(1, mPageNumber);
                                mPageNumber = 0;

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();

                                //헤더 인쇄
                                HeaderWrite_1(pCORP_NAME, pPERIOD_NAME);

                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;

                                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                                mCurrentRow = 8;       //현재 인쇄되는 row 위치.
                                mDefaultPageRow = 7;    // 페이지 증가후 PageCount 기본값.
                                mCopyLineSUM = 1;
                            }
                            else
                            {
                                mCurrentRow = mCurrentRow + (mPrintingLastRow - vPageRowCount);
                                vPageRowCount = mPrintingLastRow;

                                IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                                if (mIsNewPage == true)
                                {
                                    mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                    vPageRowCount = mDefaultPageRow;
                                }
                            }
                        }  

                        mCurrentRow = XLLine_1(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;
                         
                        vPERSON_NUM = iString.ISNull(vRow["PERSON_NUM"]); 
                    }
                }
                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }
        
        public int LineWrite_2(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty;

            bool vPrint_Flag = false;
            string vACCOUNT_GROUP_CODE = String.Empty;
            string vACCOUNT_CODE = String.Empty;

            int vLIneRow = 0;
            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vAccountRow = 4;

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 42;
            mCopy_EndRow = 33;
            mPrintingLastRow = 32;  //최종 인쇄 라인.

            mCurrentRow = 7;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 6;    // 페이지 증가후 PageCount 기본값.

            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pAdapter.OraSelectData.Rows.Count;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                #region ----- Header Write ----

                //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                //XLHeader1(pGrid, vGDColumn, vXLColumn);  // 헤더 인쇄.

                #endregion;

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                    {
                        vLIneRow++;

                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        //계정코드 동일여부 체크하여 다르면 page skip.
                        if (vACCOUNT_GROUP_CODE == null || vACCOUNT_GROUP_CODE == string.Empty)
                        {

                        }
                        else if (vACCOUNT_GROUP_CODE != iString.ISNull(vRow["ACCOUNT_GROUP_CODE"], "-"))
                        {
                            mIsNewPage = true;
                            mCurrentRow = mCurrentRow + (mPrintingLastRow - vPageRowCount);
                            vPageRowCount = mPrintingLastRow;

                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                vAccountRow = vAccountRow + mCopy_EndRow;

                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                        //계정코드 동일 여부 체크.
                        vPrint_Flag = true;
                        if (vACCOUNT_CODE == null || vACCOUNT_CODE == string.Empty || mIsNewPage == true)
                        {

                        }
                        else if (vACCOUNT_CODE != iString.ISNull(vRow["ACCOUNT_CODE"]))
                        {

                        }
                        else
                        {
                            vPrint_Flag = false;
                            mPrinting.XL_LineClearTOP(mCurrentRow, 1, 8);
                        }
                        vACCOUNT_CODE = iString.ISNull(vRow["ACCOUNT_CODE"]);

                        mCurrentRow = XLLine_2(vRow, mCurrentRow, vAccountRow, vPrint_Flag); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow - 1)
                        {
                            //XL_TOTAL_Line(12, vXLColumn);
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                vAccountRow = vAccountRow + mCopy_EndRow;

                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                        vACCOUNT_GROUP_CODE = iString.ISNull(vRow["ACCOUNT_GROUP_CODE"]);
                    }
                }
                #endregion;
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

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow + 1);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            mPageNumber++; //페이지 번호
            return vPasteEndRow;


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            //mPageNumber++; //페이지 번호
            //// 페이지 번호 표시.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //페이지 번호, XLcell[행, 열]

            //return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND,1);
        }

        public void Preview_Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrintPreview(); 
        }

        #endregion;

        #region ----- Save Methods ----

        //Excel Save//
        public void SAVE(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }

            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }
        
        //PDF Method//
        public void PDF(string pSaveFileName)
        {
            try
            {
                bool isSuccess = mPrinting.XLSaveAs_PDF(pSaveFileName); 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        } 

        #endregion;
    }
}
