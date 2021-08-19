using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0781
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
        private string mSourceSheet1 = "Source1";
        private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 - 원화 인쇄 정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 31;
        private int mCopy_EndRow = 53;
        private int mPrintingLastRow = 52;  //실제 데이터 인쇄 최종 라인.

        private int mCurrentRow = 12;        //실제 인쇄되는 row 위치.
        private int mDefaultPageRow = 11;    //페이지 skip후 적용되는 기본 PageCount 기본값.

        // 인쇄2 - 소득세 납부서 인쇄 정보.
        private int mCopy_StartCol2 = 1;
        private int mCopy_StartRow2 = 1;
        private int mCopy_EndCol2 = 33;
        private int mCopy_EndRow2 = 59;
        private int mPrintingLastRow2 = 59;  //실제 데이터 인쇄 최종 라인.

        private int mCurrentRow2 = 12;        //실제 인쇄되는 row 위치.
        private int mDefaultPageRow2 = 11;    //페이지 skip후 적용되는 기본 PageCount 기본값.

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

        #region ----- Array Set 0 ----

        private void SetArray0(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[3];
            pXLColumn = new int[3];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 12;
            pXLColumn[1] = 22;
            pXLColumn[2] = 34;
        }

        #endregion;

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("PERSON_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("NAME");
            pGDColumn[2] = pGrid.GetColumnToIndex("REPRE_NUM");
            pGDColumn[3] = pGrid.GetColumnToIndex("DEPT_NAME");
            pGDColumn[4] = pGrid.GetColumnToIndex("FLOOR_NAME");
            pGDColumn[5] = pGrid.GetColumnToIndex("ABIL_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("POST_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("ORI_JOIN_DATE");
            pGDColumn[8] = pGrid.GetColumnToIndex("JOIN_DATE");
            pGDColumn[9] = pGrid.GetColumnToIndex("RETIRE_DATE");
            pGDColumn[10] = pGrid.GetColumnToIndex("CONTINUE_YEAR");
            pGDColumn[11] = pGrid.GetColumnToIndex("END_SCH_NAME");


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 1;
            pXLColumn[1] = 6;
            pXLColumn[2] = 11;
            pXLColumn[3] = 17;
            pXLColumn[4] = 24;
            pXLColumn[5] = 31;
            pXLColumn[6] = 36;
            pXLColumn[7] = 42;
            pXLColumn[8] = 46;
            pXLColumn[9] = 50;
            pXLColumn[10] = 54;
            pXLColumn[11] = 59;
        }

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

        #region ----- Excel Write -----

        #region ----- Header Write Method ----

        public void HeaderWrite(System.Data.DataRow pRow)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;
            object vValue = null;
            string vString = string.Empty;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);
                //귀속년도
                vXLine = 3;
                vXLColumn = 4;
                vValue = pRow["WITHHOLDING_YEAR"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //내외국인-내국인
                vXLine = 3;
                vXLColumn = 40;
                vValue = pRow["NATIONALITY_1"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //내외국인-외국인
                vXLine = 4;
                vXLColumn = 40;
                vValue = pRow["NATIONALITY_9"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //거주지국
                vXLine = 5;
                vXLColumn = 31;
                vValue = pRow["NATION_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //거주지CODE
                vXLine = 5;
                vXLColumn = 38;
                vValue = pRow["NATION_ISO_CODE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자 사업자등록번호
                vXLine = 8;
                vXLColumn = 12;
                vValue = pRow["VAT_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자 법인명
                vXLine = 8;
                vXLColumn = 26;
                vValue = pRow["CORP_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자 대표자명
                vXLine = 8;
                vXLColumn = 38;
                vValue = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자 주민(법인)등록번호
                vXLine = 9;
                vXLColumn = 12;
                vValue = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자 주소
                vXLine = 9;
                vXLColumn = 26;
                vValue = pRow["CORP_ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 상호
                vXLine = 10;
                vXLColumn = 12;
                vValue = pRow["COMPANY_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 사업자등록번호
                vXLine = 10;
                vXLColumn = 33;
                vValue = pRow["TAX_REG_NO"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 사업장 주소
                vXLine = 11;
                vXLColumn = 12;
                vValue = pRow["COMPANY_ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 성명
                vXLine = 12;
                vXLColumn = 12;
                vValue = pRow["NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 주민등록번호
                vXLine = 12;
                vXLColumn = 33;
                vValue = pRow["REPRE_NUM"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득자 주소
                vXLine = 13;
                vXLColumn = 12;
                vValue = pRow["ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //업종구분
                vXLine = 14;
                vXLColumn = 7;
                vValue = pRow["BUSINESS_CODE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //인쇄일자
                vXLine = 30;
                vXLColumn = 17;
                vValue = pRow["PRINT_DATE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //징수의무자
                vXLine = 31;
                vXLColumn = 21;
                vValue = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //세무서
                vXLine = 32;
                vXLColumn = 1;
                vValue = pRow["TAX_OFFICE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 (합계) Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 0; //엑셀에 내용이 표시되는 행 번호

            int vIDX_VAT_TYPE = pGrid.GetColumnToIndex("VAT_TYPE");
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

                for (int i = 0; i < pGrid.RowCount; i++)
                {
                    // 총합계 구분에 따라 인쇄 ROW 지정.
                    if ("T" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//총합계
                        vXLine = 9;
                    }
                    else if ("3" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//신용카드.
                        vXLine = 13;
                    }
                    else if ("11" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//현금영수증.
                        vXLine = 10;
                    }

                    //0 - 거래건수.
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
                    //1 - 공급가액
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
                    //2 - 세액
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

        #region ----- Excel Write [KRW] Method -----

        private int LineWrite(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //신고구분.
                //매월
                vXLine = 3;
                vXLColumn = 2;
                vObject = pRow["MONTHLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }
                
                //반기
                vXLine = 3;
                vXLColumn = 5;
                vObject = pRow["HALF_YEARLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //수정
                vXLine = 3;
                vXLColumn = 6;
                vObject = pRow["MODIFY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //연말
                vXLine = 3;
                vXLColumn = 7;
                vObject = pRow["YEAR_END_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //소득처분
                vXLine = 3;
                vXLColumn = 9;
                vObject = pRow["INCOME_DISPOSED_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //환급신청
                vXLine = 3;
                vXLColumn = 10;
                vObject = pRow["REFUND_REQUEST_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }


                //귀속연월
                vXLine = 2;
                vXLColumn = 28;
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지급연월
                vXLine = 3;
                vXLColumn = 28;
                vObject = pRow["PAY_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인명
                vXLine = 4;
                vXLColumn = 8;
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //대표자
                vXLine = 4;
                vXLColumn = 17;
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //일괄납부여부
                vXLine = 4;
                vXLColumn = 28;
                vObject = pRow["ALL_PAYMENT_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자단위과세여부
                vXLine = 5;
                vXLColumn = 28;
                vObject = pRow["BUSINESS_UNIT_TAX_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자등록번호
                vXLine = 6;
                vXLColumn = 8;
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업장 주소
                vXLine = 6;
                vXLColumn = 17;
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전화번호
                vXLine = 6;
                vXLColumn = 27;
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이메일 주소
                vXLine = 7;
                vXLColumn = 27;
                vObject = pRow["EMAIL"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //간이세액 인원수
                vXLine = 12;
                vXLColumn = 10;
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //간이세액 총지급액
                vXLine = 12;
                vXLColumn = 12;
                vObject = pRow["A01_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //간이세액 소득세등
                vXLine = 12;
                vXLColumn = 16;
                vObject = pRow["A01_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                                
                //간이세액 농특세
                vXLine = 12;
                vXLColumn = 20;
                vObject = pRow["A01_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //간이세액 가산세
                vXLine = 12;
                vXLColumn = 23;
                vObject = pRow["A01_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 인원수
                vXLine = 13;
                vXLColumn = 10;
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 총지급액
                vXLine = 13;
                vXLColumn = 12;
                vObject = pRow["A02_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 소득세등
                vXLine = 13;
                vXLColumn = 16;
                vObject = pRow["A02_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 농특세
                vXLine = 13;
                vXLColumn = 20;
                vObject = pRow["A02_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 가산세
                vXLine = 13;
                vXLColumn = 23;
                vObject = pRow["A02_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 인원수
                vXLine = 14;
                vXLColumn = 10;
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 총지급액
                vXLine = 14;
                vXLColumn = 12;
                vObject = pRow["A03_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 소득세등
                vXLine = 14;
                vXLColumn = 16;
                vObject = pRow["A03_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 가산세
                vXLine = 14;
                vXLColumn = 23;
                vObject = pRow["A03_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 인원수
                vXLine = 15;
                vXLColumn = 10;
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 총지급액
                vXLine = 15;
                vXLColumn = 12;
                vObject = pRow["A04_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 소득세등
                vXLine = 15;
                vXLColumn = 16;
                vObject = pRow["A04_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 농특세
                vXLine = 15;
                vXLColumn = 20;
                vObject = pRow["A04_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 가산세
                vXLine = 15;
                vXLColumn = 23;
                vObject = pRow["A04_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 인원수
                vXLine = 16;
                vXLColumn = 10;
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 총지급액
                vXLine = 16;
                vXLColumn = 12;
                vObject = pRow["A10_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 소득세등
                vXLine = 16;
                vXLColumn = 16;
                vObject = pRow["A10_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 농특세
                vXLine = 16;
                vXLColumn = 20;
                vObject = pRow["A10_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 가산세
                vXLine = 16;
                vXLColumn = 23;
                vObject = pRow["A10_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 당월 조정 환급세액
                vXLine = 16;
                vXLColumn = 25;
                vObject = pRow["A10_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 16;
                vXLColumn = 27;
                vObject = pRow["A10_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 납부세액 농특세
                vXLine = 16;
                vXLColumn = 30;
                vObject = pRow["A10_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------------
                //퇴직소득 연금계좌 인원 
                vXLine = 17;
                vXLColumn = 10;
                vObject = pRow["A21_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 총지급액
                vXLine = 17;
                vXLColumn = 12;
                vObject = pRow["A21_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 소득세등
                vXLine = 17;
                vXLColumn = 16;
                vObject = pRow["A21_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 가산세
                vXLine = 17;
                vXLColumn = 23;
                vObject = pRow["A21_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 납부세액 소득세등 가산세 포함
                vXLine = 17;
                vXLColumn = 27;
                vObject = pRow["A21_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 인원 
                vXLine = 18;
                vXLColumn = 10;
                vObject = pRow["A22_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 총지급액
                vXLine = 18;
                vXLColumn = 12;
                vObject = pRow["A22_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 소득세등
                vXLine = 18;
                vXLColumn = 16;
                vObject = pRow["A22_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 가산세
                vXLine = 18;
                vXLColumn = 23;
                vObject = pRow["A22_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 납부세액 소득세등 가산세 포함
                vXLine = 18;
                vXLColumn = 27;
                vObject = pRow["A22_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //퇴직소득 인원수
                vXLine = 19;
                vXLColumn = 10;
                vObject = pRow["A20_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 총지급액
                vXLine = 19;
                vXLColumn = 12;
                vObject = pRow["A20_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 소득세등
                vXLine = 19;
                vXLColumn = 16;
                vObject = pRow["A20_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 가산세
                vXLine = 19;
                vXLColumn = 23;
                vObject = pRow["A20_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 당월 조정 환급세액
                vXLine = 19;
                vXLColumn = 25;
                vObject = pRow["A20_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 납부세액 소득세등 가산세 포함
                vXLine = 19;
                vXLColumn = 27;
                vObject = pRow["A20_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                
                //------------------------------------------------------------------------------------------
                //사업소득 매월징수 인원수
                vXLine = 20;
                vXLColumn = 10;
                vObject = pRow["A25_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 총지급액
                vXLine = 20;
                vXLColumn = 12;
                vObject = pRow["A25_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 소득세등
                vXLine = 20;
                vXLColumn = 16;
                vObject = pRow["A25_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 가산세
                vXLine = 20;
                vXLColumn = 23;
                vObject = pRow["A25_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //사업소득 연말정산 인원수
                vXLine = 21;
                vXLColumn = 10;
                vObject = pRow["A26_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 총지급액
                vXLine = 21;
                vXLColumn = 12;
                vObject = pRow["A26_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 소득세등
                vXLine = 21;
                vXLColumn = 16;
                vObject = pRow["A26_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 농특세
                vXLine = 21;
                vXLColumn = 20;
                vObject = pRow["A26_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 가산세
                vXLine = 21;
                vXLColumn = 23;
                vObject = pRow["A26_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //사업소득 가감계 인원수
                vXLine = 22;
                vXLColumn = 10;
                vObject = pRow["A30_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 총지급액
                vXLine = 22;
                vXLColumn = 12;
                vObject = pRow["A30_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 소득세등
                vXLine = 22;
                vXLColumn = 16;
                vObject = pRow["A30_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 농특세
                vXLine = 22;
                vXLColumn = 20;
                vObject = pRow["A30_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 가산세
                vXLine = 22;
                vXLColumn = 23;
                vObject = pRow["A30_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 조정 환급세액
                vXLine = 22;
                vXLColumn = 25;
                vObject = pRow["A30_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 22;
                vXLColumn = 27;
                vObject = pRow["A30_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 납부세액 농특세
                vXLine = 22;
                vXLColumn = 30;
                vObject = pRow["A30_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //------------------------------------------------------------------------------------------
                //기타소득 연금계좌 인원수
                vXLine = 23;
                vXLColumn = 10;
                vObject = pRow["A41_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 총지급액
                vXLine = 23;
                vXLColumn = 12;
                vObject = pRow["A41_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 소득세등
                vXLine = 23;
                vXLColumn = 16;
                vObject = pRow["A41_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 가산세
                vXLine = 23;
                vXLColumn = 23;
                vObject = pRow["A41_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 납부세액 소득세등 가산세 포함
                vXLine = 23;
                vXLColumn = 27;
                vObject = pRow["A41_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 인원수
                vXLine = 24;
                vXLColumn = 10;
                vObject = pRow["A42_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 총지급액
                vXLine = 24;
                vXLColumn = 12;
                vObject = pRow["A42_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 소득세등
                vXLine = 24;
                vXLColumn = 16;
                vObject = pRow["A42_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 가산세
                vXLine = 24;
                vXLColumn = 23;
                vObject = pRow["A42_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 납부세액 소득세등 가산세 포함
                vXLine = 24;
                vXLColumn = 27;
                vObject = pRow["A42_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                
                //기타소득 가감계 인원수
                vXLine = 25;
                vXLColumn = 10;
                vObject = pRow["A40_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 총지급액
                vXLine = 25;
                vXLColumn = 12;
                vObject = pRow["A40_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 소득세등
                vXLine = 25;
                vXLColumn = 16;
                vObject = pRow["A40_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 가산세
                vXLine = 25;
                vXLColumn = 23;
                vObject = pRow["A40_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 조정 환급세액
                vXLine = 25;
                vXLColumn = 25;
                vObject = pRow["A40_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 25;
                vXLColumn = 27;
                vObject = pRow["A40_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //------------------------------------------------------------------------------------------
                //연금소득 연금계좌 인원수
                vXLine = 26;
                vXLColumn = 10;
                vObject = pRow["A48_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 총지급액
                vXLine = 26;
                vXLColumn = 12;
                vObject = pRow["A48_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 소득세등
                vXLine = 26;
                vXLColumn = 16;
                vObject = pRow["A48_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 가산세
                vXLine = 26;
                vXLColumn = 23;
                vObject = pRow["A48_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 인원수
                vXLine = 27;
                vXLColumn = 10;
                vObject = pRow["A45_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 총지급액
                vXLine = 27;
                vXLColumn = 12;
                vObject = pRow["A45_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 소득세등
                vXLine = 27;
                vXLColumn = 16;
                vObject = pRow["A45_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 가산세
                vXLine = 27;
                vXLColumn = 23;
                vObject = pRow["A45_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //연금소득 연말정산 인원수
                vXLine = 28;
                vXLColumn = 10;
                vObject = pRow["A46_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 총지급액
                vXLine = 28;
                vXLColumn = 12;
                vObject = pRow["A46_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 소득세등
                vXLine = 28;
                vXLColumn = 16;
                vObject = pRow["A46_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 가산세
                vXLine = 28;
                vXLColumn = 23;
                vObject = pRow["A46_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                
                //연금소득 가감계 인원수
                vXLine = 29;
                vXLColumn = 10;
                vObject = pRow["A47_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 총지급액
                vXLine = 29;
                vXLColumn = 12;
                vObject = pRow["A47_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 소득세등
                vXLine = 29;
                vXLColumn = 16;
                vObject = pRow["A47_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 가산세
                vXLine = 29;
                vXLColumn = 23;
                vObject = pRow["A47_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 조정 환급세액
                vXLine = 29;
                vXLColumn = 25;
                vObject = pRow["A47_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 29;
                vXLColumn = 27;
                vObject = pRow["A47_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------
                //이자소득 인원수
                vXLine = 30;
                vXLColumn = 10;
                vObject = pRow["A50_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 총지급액
                vXLine = 30;
                vXLColumn = 12;
                vObject = pRow["A50_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 소득세등
                vXLine = 30;
                vXLColumn = 16;
                vObject = pRow["A50_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 농특세
                vXLine = 30;
                vXLColumn = 20;
                vObject = pRow["A50_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                
                //이자소득 가산세
                vXLine = 30;
                vXLColumn = 23;
                vObject = pRow["A50_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 조정 환급세액
                vXLine = 30;
                vXLColumn = 25;
                vObject = pRow["A50_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 납부세액 소득세등 가산세 포함
                vXLine = 30;
                vXLColumn = 27;
                vObject = pRow["A50_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 납부세액 농특세
                vXLine = 30;
                vXLColumn = 30;
                vObject = pRow["A50_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //배당소득 인원수
                vXLine = 31;
                vXLColumn = 10;
                vObject = pRow["A60_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 총지급액
                vXLine = 31;
                vXLColumn = 12;
                vObject = pRow["A60_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 소득세등
                vXLine = 31;
                vXLColumn = 16;
                vObject = pRow["A60_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 농특세
                vXLine = 31;
                vXLColumn = 20;
                vObject = pRow["A60_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 가산세
                vXLine = 31;
                vXLColumn = 23;
                vObject = pRow["A60_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 조정 환급세액
                vXLine = 31;
                vXLColumn = 25;
                vObject = pRow["A60_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 납부세액 소득세등 가산세 포함
                vXLine = 31;
                vXLColumn = 27;
                vObject = pRow["A60_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 납부세액 농특세
                vXLine = 31;
                vXLColumn = 30;
                vObject = pRow["A60_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------------
                //저축해지 인원수
                vXLine = 32;
                vXLColumn = 10;
                vObject = pRow["A69_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 소득세등
                vXLine = 32;
                vXLColumn = 16;
                vObject = pRow["A69_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 가산세
                vXLine = 32;
                vXLColumn = 23;
                vObject = pRow["A69_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 조정 환급세액
                vXLine = 32;
                vXLColumn = 25;
                vObject = pRow["A69_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 납부세액 소득세등 가산세 포함
                vXLine = 32;
                vXLColumn = 27;
                vObject = pRow["A69_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //비거주자 인원수
                vXLine = 33;
                vXLColumn = 10;
                vObject = pRow["A70_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 총지급액
                vXLine = 33;
                vXLColumn = 12;
                vObject = pRow["A70_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 소득세등
                vXLine = 33;
                vXLColumn = 16;
                vObject = pRow["A70_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 가산세
                vXLine = 33;
                vXLColumn = 23;
                vObject = pRow["A70_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 조정 환급세액
                vXLine = 33;
                vXLColumn = 25;
                vObject = pRow["A70_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 납부세액 소득세등 가산세 포함
                vXLine = 33;
                vXLColumn = 27;
                vObject = pRow["A70_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //내외국인법인원천 인원수
                vXLine = 34;
                vXLColumn = 10;
                vObject = pRow["A80_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 총지급액
                vXLine = 34;
                vXLColumn = 12;
                vObject = pRow["A80_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 소득세등
                vXLine = 34;
                vXLColumn = 16;
                vObject = pRow["A80_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 가산세
                vXLine = 34;
                vXLColumn = 23;
                vObject = pRow["A80_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 조정 환급세액
                vXLine = 34;
                vXLColumn = 25;
                vObject = pRow["A80_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 납부세액 소득세등 가산세 포함
                vXLine = 34;
                vXLColumn = 27;
                vObject = pRow["A80_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //---------------------------------------------------------------------
                //수정신고 소득세등
                vXLine = 35;
                vXLColumn = 16;
                vObject = pRow["A90_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 농특세
                vXLine = 35;
                vXLColumn = 20;
                vObject = pRow["A90_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 가산세
                vXLine = 35;
                vXLColumn = 23;
                vObject = pRow["A90_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 조정 환급세액
                vXLine = 35;
                vXLColumn = 25;
                vObject = pRow["A90_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 납부세액 소득세등 가산세 포함
                vXLine = 35;
                vXLColumn = 27;
                vObject = pRow["A90_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 납부세액 농특세
                vXLine = 35;
                vXLColumn = 30;
                vObject = pRow["A90_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //---------------------------------------------------------------------------------
                //총합계 인원수
                vXLine = 36;
                vXLColumn = 10;
                vObject = pRow["A99_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 총지급액
                vXLine = 36;
                vXLColumn = 12;
                vObject = pRow["A99_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 소득세등
                vXLine = 36;
                vXLColumn = 16;
                vObject = pRow["A99_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 농특세
                vXLine = 36;
                vXLColumn = 20;
                vObject = pRow["A99_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 가산세
                vXLine = 36;
                vXLColumn = 23;
                vObject = pRow["A99_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 조정 환급세액
                vXLine = 36;
                vXLColumn = 25;
                vObject = pRow["A99_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 납부세액 소득세등 가산세 포함
                vXLine = 36;
                vXLColumn = 27;
                vObject = pRow["A99_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 납부세액 농특세
                vXLine = 36;
                vXLColumn = 30;
                vObject = pRow["A99_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //--------------------------------------------------------------------------------
                //12.전월미환급세액
                vXLine = 41;
                vXLColumn = 2;
                vObject = pRow["RECEIVE_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //13.기환급신청한세액
                vXLine = 41;
                vXLColumn = 6;
                vObject = pRow["ALREADY_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //14.차감잔액
                vXLine = 41;
                vXLColumn = 9;
                vObject = pRow["REFUND_BALANCE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //15.일반환급
                vXLine = 41;
                vXLColumn = 12;
                vObject = pRow["GENERAL_REFUND_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //16.신탁재산(금융회사등)
                vXLine = 41;
                vXLColumn = 15;
                vObject = pRow["FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //17-1.그밖의 환급세액-금융회사등
                vXLine = 41;
                vXLColumn = 18;
                vObject = pRow["ETC_REFUND_FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //17-2.그밖의 환급세액-합병등
                vXLine = 41;
                vXLColumn = 20;
                vObject = pRow["ETC_REFUND_MERGER_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                
                //18.조정대상환급세액
                vXLine = 41;
                vXLColumn = 22;
                vObject = pRow["ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));
                
                //19. 당월 조정정환급세액
                vXLine = 41;
                vXLColumn = 25;
                vObject = pRow["THIS_ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //20. 차월 이월 환급세액
                vXLine = 41;
                vXLColumn = 27;
                vObject = pRow["NEXT_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //21.환급신청액
                vXLine = 41;
                vXLColumn = 30;
                vObject = pRow["REQUEST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //--------------------------------------------------------------
                //제출일자
                vXLine = 45;
                vXLColumn = 10;
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신고인
                vXLine = 46;
                vXLColumn = 7;
                vObject = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //세무서
                vXLine = 52;
                vXLColumn = 3;
                vObject = pRow["TAX_OFFICE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite_11(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //타이틀
                vXLine = 2;
                vXLColumn = 12;                
                if (iString.ISNull(pRow["REQUEST_REFUND_FLAG"]) == "Y")
                {
                    vObject = "□ 원천징수이행상황신고서\r\n■ 원천징수세액환급신청서";
                }
                else
                {
                    vObject = "■ 원천징수이행상황신고서\r\n□ 원천징수세액환급신청서";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //신고구분.
                //매월
                vXLine = 3;
                vXLColumn = 2;
                vObject = pRow["MONTHLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //반기
                vXLine = 3;
                vXLColumn = 5;
                vObject = pRow["HALF_YEARLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //수정
                vXLine = 3;
                vXLColumn = 6;
                vObject = pRow["MODIFY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //연말
                vXLine = 3;
                vXLColumn = 7;
                vObject = pRow["YEAR_END_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //소득처분
                vXLine = 3;
                vXLColumn = 9;
                vObject = pRow["INCOME_DISPOSED_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //환급신청
                vXLine = 3;
                vXLColumn = 10;
                vObject = pRow["REFUND_REQUEST_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }


                //귀속연월
                vXLine = 2;
                vXLColumn = 28;
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지급연월
                vXLine = 3;
                vXLColumn = 28;
                vObject = pRow["PAY_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인명
                vXLine = 4;
                vXLColumn = 8;
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //대표자
                vXLine = 4;
                vXLColumn = 17;
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //일괄납부여부
                vXLine = 4;
                vXLColumn = 28;
                vObject = pRow["ALL_PAYMENT_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자단위과세여부
                vXLine = 5;
                vXLColumn = 28;
                vObject = pRow["BUSINESS_UNIT_TAX_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자등록번호
                vXLine = 6;
                vXLColumn = 8;
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업장 주소
                vXLine = 6;
                vXLColumn = 17;
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전화번호
                vXLine = 6;
                vXLColumn = 27;
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이메일 주소
                vXLine = 7;
                vXLColumn = 27;
                vObject = pRow["EMAIL"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //간이세액 인원수
                vXLine = 12;
                vXLColumn = 10;
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //간이세액 총지급액
                vXLine = 12;
                vXLColumn = 12;
                vObject = pRow["A01_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //간이세액 소득세등
                vXLine = 12;
                vXLColumn = 16;
                vObject = pRow["A01_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //간이세액 농특세
                vXLine = 12;
                vXLColumn = 20;
                vObject = pRow["A01_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //간이세액 가산세
                vXLine = 12;
                vXLColumn = 23;
                vObject = pRow["A01_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 인원수
                vXLine = 13;
                vXLColumn = 10;
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 총지급액
                vXLine = 13;
                vXLColumn = 12;
                vObject = pRow["A02_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 소득세등
                vXLine = 13;
                vXLColumn = 16;
                vObject = pRow["A02_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 농특세
                vXLine = 13;
                vXLColumn = 20;
                vObject = pRow["A02_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //중도퇴사 가산세
                vXLine = 13;
                vXLColumn = 23;
                vObject = pRow["A02_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 인원수
                vXLine = 14;
                vXLColumn = 10;
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 총지급액
                vXLine = 14;
                vXLColumn = 12;
                vObject = pRow["A03_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 소득세등
                vXLine = 14;
                vXLColumn = 16;
                vObject = pRow["A03_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //일용근로 가산세
                vXLine = 14;
                vXLColumn = 23;
                vObject = pRow["A03_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 합계 
                vXLine = 15;
                vXLColumn = 10;
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 합계 총지급액
                vXLine = 15;
                vXLColumn = 12;
                vObject = pRow["A04_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 합계 소득세등
                vXLine = 15;
                vXLColumn = 16;
                vObject = pRow["A04_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 합계 농특세
                vXLine = 15;
                vXLColumn = 20;
                vObject = pRow["A04_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 합계 가산세
                vXLine = 15;
                vXLColumn = 23;
                vObject = pRow["A04_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 분납 신청자 인원수
                vXLine = 16;
                vXLColumn = 10;
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 분납 총지급액
                vXLine = 16;
                vXLColumn = 12;
                vObject = pRow["A05_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 분납 소득세등
                vXLine = 16;
                vXLColumn = 16;
                vObject = pRow["A05_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 분납 농특세
                vXLine = 16;
                vXLColumn = 20;
                vObject = pRow["A05_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 분납 가산세
                vXLine = 16;
                vXLColumn = 23;
                vObject = pRow["A05_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 납부 인원수
                vXLine = 17;
                vXLColumn = 10;
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 납부 총지급액
                vXLine = 17;
                vXLColumn = 12;
                vObject = pRow["A06_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 납부 소득세등
                vXLine = 17;
                vXLColumn = 16;
                vObject = pRow["A06_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 납부 농특세
                vXLine = 17;
                vXLColumn = 20;
                vObject = pRow["A06_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연말정산 납부 가산세
                vXLine = 17;
                vXLColumn = 23;
                vObject = pRow["A06_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 인원수
                vXLine = 18;
                vXLColumn = 10;
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 총지급액
                vXLine = 18;
                vXLColumn = 12;
                vObject = pRow["A10_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 소득세등
                vXLine = 18;
                vXLColumn = 16;
                vObject = pRow["A10_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 농특세
                vXLine = 18;
                vXLColumn = 20;
                vObject = pRow["A10_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 가산세
                vXLine = 18;
                vXLColumn = 23;
                vObject = pRow["A10_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 당월 조정 환급세액
                vXLine = 18;
                vXLColumn = 25;
                vObject = pRow["A10_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 18;
                vXLColumn = 27;
                vObject = pRow["A10_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //근로소득 가감계 납부세액 농특세
                vXLine = 18;
                vXLColumn = 30;
                vObject = pRow["A10_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------------
                //퇴직소득 연금계좌 인원 
                vXLine = 19;
                vXLColumn = 10;
                vObject = pRow["A21_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 총지급액
                vXLine = 19;
                vXLColumn = 12;
                vObject = pRow["A21_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 소득세등
                vXLine = 19;
                vXLColumn = 16;
                vObject = pRow["A21_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 가산세
                vXLine = 19;
                vXLColumn = 23;
                vObject = pRow["A21_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 연금계좌 납부세액 소득세등 가산세 포함
                vXLine = 19;
                vXLColumn = 27;
                vObject = pRow["A21_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 인원 
                vXLine = 20;
                vXLColumn = 10;
                vObject = pRow["A22_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 총지급액
                vXLine = 20;
                vXLColumn = 12;
                vObject = pRow["A22_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 소득세등
                vXLine = 20;
                vXLColumn = 16;
                vObject = pRow["A22_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 가산세
                vXLine = 20;
                vXLColumn = 23;
                vObject = pRow["A22_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 그외 납부세액 소득세등 가산세 포함
                vXLine = 20;
                vXLColumn = 27;
                vObject = pRow["A22_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //퇴직소득 인원수
                vXLine = 21;
                vXLColumn = 10;
                vObject = pRow["A20_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 총지급액
                vXLine = 21;
                vXLColumn = 12;
                vObject = pRow["A20_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 소득세등
                vXLine = 21;
                vXLColumn = 16;
                vObject = pRow["A20_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 가산세
                vXLine = 21;
                vXLColumn = 23;
                vObject = pRow["A20_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 당월 조정 환급세액
                vXLine = 21;
                vXLColumn = 25;
                vObject = pRow["A20_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //퇴직소득 납부세액 소득세등 가산세 포함
                vXLine = 21;
                vXLColumn = 27;
                vObject = pRow["A20_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //------------------------------------------------------------------------------------------
                //사업소득 매월징수 인원수
                vXLine = 22;
                vXLColumn = 10;
                vObject = pRow["A25_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 총지급액
                vXLine = 22;
                vXLColumn = 12;
                vObject = pRow["A25_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 소득세등
                vXLine = 22;
                vXLColumn = 16;
                vObject = pRow["A25_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 매월징수 가산세
                vXLine = 22;
                vXLColumn = 23;
                vObject = pRow["A25_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //사업소득 연말정산 인원수
                vXLine = 23;
                vXLColumn = 10;
                vObject = pRow["A26_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 총지급액
                vXLine = 23;
                vXLColumn = 12;
                vObject = pRow["A26_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 소득세등
                vXLine = 23;
                vXLColumn = 16;
                vObject = pRow["A26_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 농특세
                vXLine = 23;
                vXLColumn = 20;
                vObject = pRow["A26_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 연말정산 가산세
                vXLine = 23;
                vXLColumn = 23;
                vObject = pRow["A26_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //사업소득 가감계 인원수
                vXLine = 24;
                vXLColumn = 10;
                vObject = pRow["A30_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 총지급액
                vXLine = 24;
                vXLColumn = 12;
                vObject = pRow["A30_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 소득세등
                vXLine = 24;
                vXLColumn = 16;
                vObject = pRow["A30_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 농특세
                vXLine = 24;
                vXLColumn = 20;
                vObject = pRow["A30_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 가산세
                vXLine = 24;
                vXLColumn = 23;
                vObject = pRow["A30_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 조정 환급세액
                vXLine = 24;
                vXLColumn = 25;
                vObject = pRow["A30_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 24;
                vXLColumn = 27;
                vObject = pRow["A30_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //사업소득 가감계 납부세액 농특세
                vXLine = 24;
                vXLColumn = 30;
                vObject = pRow["A30_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //----------------------------------------2017 개정추가

                //기타소득 연금계좌 인원수
                vXLine = 25;
                vXLColumn = 10;
                vObject = pRow["A41_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 총지급액
                vXLine = 25;
                vXLColumn = 12;
                vObject = pRow["A41_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 소득세등
                vXLine = 25;
                vXLColumn = 16;
                vObject = pRow["A41_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 가산세
                vXLine = 25;
                vXLColumn = 23;
                vObject = pRow["A41_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 연금계좌 납부세액 소득세등 가산세 포함
                vXLine = 25;
                vXLColumn = 27;
                vObject = pRow["A41_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                
                ////기타소득 종교인소득 매월징수 인원수
                vXLine = 26;
                vXLColumn = 10;
                vObject = pRow["A43_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득 종교인소득 매월징수 총지급액
                vXLine = 26;
                vXLColumn = 12;
                vObject = pRow["A43_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득 종교인소득 매월징수 소득세등
                vXLine = 26;
                vXLColumn = 16;
                vObject = pRow["A43_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득 종교인소득 매월징수 가산세
                vXLine = 26;
                vXLColumn = 23;
                vObject = pRow["A43_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득 종교인소득 매월징수 납부세액 소득세등 가산세 포함
                vXLine = 26;
                vXLColumn = 27;
                vObject = pRow["A43_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득  종교인소득 연말정산 인원수
                vXLine = 27;
                vXLColumn = 10;
                vObject = pRow["A44_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득  종교인소득 연말정산 총지급액
                vXLine = 27;
                vXLColumn = 12;
                vObject = pRow["A44_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득  종교인소득 연말정산 소득세등
                vXLine = 27;
                vXLColumn = 16;
                vObject = pRow["A44_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득  종교인소득 연말정산 가산세
                vXLine = 27;
                vXLColumn = 23;
                vObject = pRow["A44_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                ////기타소득 종교인소득 연말정산 납부세액 소득세등 가산세 포함
                vXLine = 27;
                vXLColumn = 27;
                vObject = pRow["A44_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));



                //----------------------------------------2017 개정추가


                //기타소득 그외 인원수
                vXLine = 26+2;
                vXLColumn = 10;
                vObject = pRow["A42_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 총지급액
                vXLine = 26 + 2;
                vXLColumn = 12;
                vObject = pRow["A42_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 소득세등
                vXLine = 26 + 2;
                vXLColumn = 16;
                vObject = pRow["A42_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 가산세
                vXLine = 26 + 2;
                vXLColumn = 23;
                vObject = pRow["A42_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 그외 납부세액 소득세등 가산세 포함
                vXLine = 26 + 2;
                vXLColumn = 27;
                vObject = pRow["A42_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 인원수
                vXLine = 27 + 2;
                vXLColumn = 10;
                vObject = pRow["A40_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 총지급액
                vXLine = 27 + 2;
                vXLColumn = 12;
                vObject = pRow["A40_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 소득세등
                vXLine = 27 + 2;
                vXLColumn = 16;
                vObject = pRow["A40_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 가산세
                vXLine = 27 + 2;
                vXLColumn = 23;
                vObject = pRow["A40_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 조정 환급세액
                vXLine = 27 + 2;
                vXLColumn = 25;
                vObject = pRow["A40_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //기타소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 27 + 2;
                vXLColumn = 27;
                vObject = pRow["A40_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //------------------------------------------------------------------------------------------
                //연금소득 연금계좌 인원수
                vXLine = 28 + 2;
                vXLColumn = 10;
                vObject = pRow["A48_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 총지급액
                vXLine = 28 + 2;
                vXLColumn = 12;
                vObject = pRow["A48_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 소득세등
                vXLine = 28 + 2;
                vXLColumn = 16;
                vObject = pRow["A48_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연금계좌 가산세
                vXLine = 28 + 2;
                vXLColumn = 23;
                vObject = pRow["A48_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 인원수
                vXLine = 29 + 2;
                vXLColumn = 10;
                vObject = pRow["A45_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 총지급액
                vXLine = 29 + 2;
                vXLColumn = 12;
                vObject = pRow["A45_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 소득세등
                vXLine = 29 + 2;
                vXLColumn = 16;
                vObject = pRow["A45_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 공적연금(매월) 가산세
                vXLine = 29 + 2;
                vXLColumn = 23;
                vObject = pRow["A45_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));


                //연금소득 연말정산 인원수
                vXLine = 30 + 2;
                vXLColumn = 10;
                vObject = pRow["A46_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 총지급액
                vXLine = 30 + 2;
                vXLColumn = 12;
                vObject = pRow["A46_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 소득세등
                vXLine = 30 + 2;
                vXLColumn = 16;
                vObject = pRow["A46_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 연말정산 가산세
                vXLine = 30 + 2;
                vXLColumn = 23;
                vObject = pRow["A46_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 인원수
                vXLine = 31 + 2;
                vXLColumn = 10;
                vObject = pRow["A47_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 총지급액
                vXLine = 31 + 2;
                vXLColumn = 12;
                vObject = pRow["A47_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 소득세등
                vXLine = 31 + 2;
                vXLColumn = 16;
                vObject = pRow["A47_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 가산세
                vXLine = 31 + 2;
                vXLColumn = 23;
                vObject = pRow["A47_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 조정 환급세액
                vXLine = 31 + 2;
                vXLColumn = 25;
                vObject = pRow["A47_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //연금소득 가감계 납부세액 소득세등 가산세 포함
                vXLine = 31 + 2;
                vXLColumn = 27;
                vObject = pRow["A47_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------
                //이자소득 인원수
                vXLine = 32 + 2;
                vXLColumn = 10;
                vObject = pRow["A50_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 총지급액
                vXLine = 32 + 2;
                vXLColumn = 12;
                vObject = pRow["A50_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 소득세등
                vXLine = 32 + 2;
                vXLColumn = 16;
                vObject = pRow["A50_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 농특세
                vXLine = 32 + 2;
                vXLColumn = 20;
                vObject = pRow["A50_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 가산세
                vXLine = 32 + 2;
                vXLColumn = 23;
                vObject = pRow["A50_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 조정 환급세액
                vXLine = 32 + 2;
                vXLColumn = 25;
                vObject = pRow["A50_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 납부세액 소득세등 가산세 포함
                vXLine = 32 + 2;
                vXLColumn = 27;
                vObject = pRow["A50_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //이자소득 납부세액 농특세
                vXLine = 32 + 2;
                vXLColumn = 30;
                vObject = pRow["A50_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //배당소득 인원수
                vXLine = 33 + 2;
                vXLColumn = 10;
                vObject = pRow["A60_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 총지급액
                vXLine = 33 + 2;
                vXLColumn = 12;
                vObject = pRow["A60_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 소득세등
                vXLine = 33 + 2;
                vXLColumn = 16;
                vObject = pRow["A60_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 농특세
                vXLine = 33 + 2;
                vXLColumn = 20;
                vObject = pRow["A60_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 가산세
                vXLine = 33 + 2;
                vXLColumn = 23;
                vObject = pRow["A60_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 조정 환급세액
                vXLine = 33 + 2;
                vXLColumn = 25;
                vObject = pRow["A60_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 납부세액 소득세등 가산세 포함
                vXLine = 33 + 2;
                vXLColumn = 27;
                vObject = pRow["A60_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //배당소득 납부세액 농특세
                vXLine = 33 + 2;
                vXLColumn = 30;
                vObject = pRow["A60_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //----------------------------------------------------------------------------------
                //저축해지 인원수
                vXLine = 34 + 2;
                vXLColumn = 10;
                vObject = pRow["A69_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 소득세등
                vXLine = 34 + 2;
                vXLColumn = 16;
                vObject = pRow["A69_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 가산세
                vXLine = 34 + 2;
                vXLColumn = 23;
                vObject = pRow["A69_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 조정 환급세액
                vXLine = 34 + 2;
                vXLColumn = 25;
                vObject = pRow["A69_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //저축해지 납부세액 소득세등 가산세 포함
                vXLine = 34 + 2;
                vXLColumn = 27;
                vObject = pRow["A69_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //비거주자 인원수
                vXLine = 35 + 2;
                vXLColumn = 10;
                vObject = pRow["A70_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 총지급액
                vXLine = 35 + 2;
                vXLColumn = 12;
                vObject = pRow["A70_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 소득세등
                vXLine = 35 + 2;
                vXLColumn = 16;
                vObject = pRow["A70_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 가산세
                vXLine = 35 + 2;
                vXLColumn = 23;
                vObject = pRow["A70_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 조정 환급세액
                vXLine = 35 + 2;
                vXLColumn = 25;
                vObject = pRow["A70_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //비거주자 납부세액 소득세등 가산세 포함
                vXLine = 35 + 2;
                vXLColumn = 27;
                vObject = pRow["A70_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------------------
                //내외국인법인원천 인원수
                vXLine = 36 + 2;
                vXLColumn = 10;
                vObject = pRow["A80_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 총지급액
                vXLine = 36 + 2;
                vXLColumn = 12;
                vObject = pRow["A80_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 소득세등
                vXLine = 36 + 2;
                vXLColumn = 16;
                vObject = pRow["A80_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 가산세
                vXLine = 36 + 2;
                vXLColumn = 23;
                vObject = pRow["A80_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 조정 환급세액
                vXLine = 36 + 2;
                vXLColumn = 25;
                vObject = pRow["A80_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //내외국인법인원천 납부세액 소득세등 가산세 포함
                vXLine = 36 + 2;
                vXLColumn = 27;
                vObject = pRow["A80_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //---------------------------------------------------------------------
                //수정신고 소득세등
                vXLine = 37 + 2;
                vXLColumn = 16;
                vObject = pRow["A90_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 농특세
                vXLine = 37 + 2;
                vXLColumn = 20;
                vObject = pRow["A90_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 가산세
                vXLine = 37 + 2;
                vXLColumn = 23;
                vObject = pRow["A90_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 조정 환급세액
                vXLine = 37 + 2;
                vXLColumn = 25;
                vObject = pRow["A90_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 납부세액 소득세등 가산세 포함
                vXLine = 37 + 2;
                vXLColumn = 27;
                vObject = pRow["A90_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //수정신고 납부세액 농특세
                vXLine = 37 + 2;
                vXLColumn = 30;
                vObject = pRow["A90_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //---------------------------------------------------------------------------------
                //총합계 인원수
                vXLine = 38 + 2;
                vXLColumn = 10;
                vObject = pRow["A99_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 총지급액
                vXLine = 38 + 2;
                vXLColumn = 12;
                vObject = pRow["A99_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 소득세등
                vXLine = 38 + 2;
                vXLColumn = 16;
                vObject = pRow["A99_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 농특세
                vXLine = 38 + 2;
                vXLColumn = 20;
                vObject = pRow["A99_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 가산세
                vXLine = 38 + 2;
                vXLColumn = 23;
                vObject = pRow["A99_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 조정 환급세액
                vXLine = 38 + 2;
                vXLColumn = 25;
                vObject = pRow["A99_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 납부세액 소득세등 가산세 포함
                vXLine = 38 + 2;
                vXLColumn = 27;
                vObject = pRow["A99_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //총합계 납부세액 농특세
                vXLine = 38 + 2;
                vXLColumn = 30;
                vObject = pRow["A99_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //--------------------------------------------------------------------------------
                //12.전월미환급세액
                vXLine = 43 + 2;
                vXLColumn = 2;
                vObject = pRow["RECEIVE_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //13.기환급신청한세액
                vXLine = 43 + 2;
                vXLColumn = 6;
                vObject = pRow["ALREADY_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //14.차감잔액
                vXLine = 43 + 2;
                vXLColumn = 9;
                vObject = pRow["REFUND_BALANCE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //15.일반환급
                vXLine = 43 + 2;
                vXLColumn = 12;
                vObject = pRow["GENERAL_REFUND_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //16.신탁재산(금융회사등)
                vXLine = 43 + 2;
                vXLColumn = 15;
                vObject = pRow["FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //17-1.그밖의 환급세액-금융회사등
                vXLine = 43 + 2;
                vXLColumn = 18;
                vObject = pRow["ETC_REFUND_FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //17-2.그밖의 환급세액-합병등
                vXLine = 43 + 2;
                vXLColumn = 20;
                vObject = pRow["ETC_REFUND_MERGER_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //18.조정대상환급세액
                vXLine = 43 + 2;
                vXLColumn = 22;
                vObject = pRow["ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //19. 당월 조정정환급세액
                vXLine = 43 + 2;
                vXLColumn = 25;
                vObject = pRow["THIS_ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //20. 차월 이월 환급세액
                vXLine = 43 + 2;
                vXLColumn = 27;
                vObject = pRow["NEXT_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //21.환급신청액
                vXLine = 43 + 2;
                vXLColumn = 30;
                vObject = pRow["REQUEST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //--------------------------------------------------------------
                //제출일자
                vXLine = 47 + 2;
                vXLColumn = 10;
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신고인
                vXLine = 48 + 2;
                vXLColumn = 7;
                vObject = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //세무서
                vXLine = 54 + 2;
                vXLColumn = 3;
                vObject = pRow["TAX_OFFICE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }


        private int LineWrite_11_SUB_01(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //코드 
                vXLColumn = 12;
                if (iString.ISNull(pRow["INCOME_SUB_CODE"]) != string.Empty)
                {
                    vObject = string.Format("{0}", pRow["INCOME_SUB_CODE"]);
                }
                else
                {
                    vObject = "";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //인원
                vXLColumn = 14; 
                vObject = pRow["PERSON_CNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //총지급액
                vXLColumn = 16;
                vObject = pRow["PAYMENT_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세등
                vXLColumn = 19;
                vObject = pRow["INCOME_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세
                vXLColumn = 22;
                vObject = pRow["SP_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //가산세 
                vXLColumn = 24;
                vObject = pRow["ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //조정환급세액. 
                vXLColumn = 26;
                vObject = pRow["REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부 소득세  
                vXLColumn = 28;
                vObject = pRow["FIX_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //납부농특세 
                vXLColumn = 30;
                vObject = pRow["FIX_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));  

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite_11_SUB_02(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //코드 
                vXLColumn = 12;
                if (iString.ISNull(pRow["INCOME_SUB_CODE"]) != string.Empty)
                {
                    vObject = string.Format("{0}", pRow["INCOME_SUB_CODE"]);
                }
                else
                {
                    vObject = "";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //인원
                vXLColumn = 14;
                vObject = pRow["PERSON_CNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //총지급액
                vXLColumn = 16;
                vObject = pRow["PAYMENT_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세등
                vXLColumn = 19;
                vObject = pRow["INCOME_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세
                vXLColumn = 22;
                vObject = pRow["SP_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //가산세 
                vXLColumn = 24;
                vObject = pRow["ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //조정환급세액. 
                vXLColumn = 26;
                vObject = pRow["REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부 소득세  
                vXLColumn = 28;
                vObject = pRow["FIX_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //납부농특세 
                vXLColumn = 30;
                vObject = pRow["FIX_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "△"));

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- Excel Write [CURRENCY] Method -----

        private int LineWrite2(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
               //분류기호                
                vObject = pRow["CLASSIFY_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 1;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //시코드
                vObject = pRow["CITY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부년월
                vObject = pRow["SUBMIT_YYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부구분.
                vObject = pRow["SUBMIT_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //세목
                vObject = pRow["TAX_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //수입징수관서.
                vXLColumn = 22;
                vObject = pRow["TAX_OFFICE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 19;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계좌번호.
                vObject = pRow["TAX_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 23;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //상호.
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자번호.
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //회계연도.
                vObject = pRow["FISCAL_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업장주소.
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 48;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전화.
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                vXLine = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 48;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //귀속연도.
                vObject = pRow["STD_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //귀속월.
                vObject = pRow["STD_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부기한.
                vObject = pRow["PAYMENT_DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-조.
                vObject = pRow["INCOME_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-천.
                vObject = pRow["INCOME_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-백.
                vObject = pRow["INCOME_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-십.
                vObject = pRow["INCOME_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-억.
                vObject = pRow["INCOME_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-천.
                vObject = pRow["INCOME_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-백.
                vObject = pRow["INCOME_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-십.
                vObject = pRow["INCOME_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-만.
                vObject = pRow["INCOME_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-천.
                vObject = pRow["INCOME_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-백.
                vObject = pRow["INCOME_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-십.
                vObject = pRow["INCOME_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득계-일.
                vObject = pRow["INCOME_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-조.
                vObject = pRow["SP_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-천.
                vObject = pRow["SP_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-백.
                vObject = pRow["SP_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-십.
                vObject = pRow["SP_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-억.
                vObject = pRow["SP_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-천.
                vObject = pRow["SP_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-백.
                vObject = pRow["SP_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-십.
                vObject = pRow["SP_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-만.
                vObject = pRow["SP_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-천.
                vObject = pRow["SP_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-백.
                vObject = pRow["SP_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-십.
                vObject = pRow["SP_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농어촌특별세계-일.
                vObject = pRow["SP_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-조.
                vObject = pRow["SUM_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-천.
                vObject = pRow["SUM_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-백.
                vObject = pRow["SUM_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-십.
                vObject = pRow["SUM_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-억.
                vObject = pRow["SUM_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-천.
                vObject = pRow["SUM_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-백.
                vObject = pRow["SUM_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-십.
                vObject = pRow["SUM_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-만.
                vObject = pRow["SUM_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-천.
                vObject = pRow["SUM_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-백.
                vObject = pRow["SUM_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-십.
                vObject = pRow["SUM_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계-일.
                vObject = pRow["SUM_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XLTOTAL_Line(int pXLine)
        //{// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
        //    int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // 원본을 복사해서 타겟 에 복사해 넣음.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //차변합계
        //        vXLColumnIndex = 14;
        //        IsConvert = IsConvertNumber(mTOT_DR_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //        }
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

        //        //대변합계
        //        vXLColumnIndex = 20;
        //        IsConvert = IsConvertNumber(mTOT_CR_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //        }
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

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
        //    return vXLine;
        //}

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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte MAIN Methods ----

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // 인쇄 - 원화 인쇄 정보.
            mCopy_EndCol = 31;
            mCopy_EndRow = 53;
            mPrintingLastRow = 52;  //실제 데이터 인쇄 최종 라인.

            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //라인 열수.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        
                        mCurrentRow = LineWrite(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
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


        //2016년02월 변경분//
        public int ExcelWrite_11(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // 인쇄 - 원화 인쇄 정보.
            mCopy_EndCol = 31;
            mCopy_EndRow = 57;
            mPrintingLastRow = 56;  //실제 데이터 인쇄 최종 라인.

            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //라인 열수.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
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


        //2016년02월 변경분//
        public int ExcelWrite_11_SUB(InfoSummit.Win.ControlAdv.ISDataAdapter pINCOME_SUB_01, InfoSummit.Win.ControlAdv.ISDataAdapter pINCOME_SUB_02)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // 인쇄 - 원화 인쇄 정보.
            mCopy_EndCol = 31;
            mCopy_EndRow = 86;
            mPrintingLastRow = 56;  //실제 데이터 인쇄 최종 라인.

            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pINCOME_SUB_01.CurrentRows.Count;
                vTotalRow = vTotalRow + pINCOME_SUB_02.CurrentRows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste_SUB(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                    mCurrentRow = 65;
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pINCOME_SUB_01.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11_SUB_01(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
                        }
                        else
                        {
                             
                        }
                    }

                    mCurrentRow = 112;
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pINCOME_SUB_02.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11_SUB_02(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
                        }
                        else
                        {

                        }
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

        #region ----- 소득세납부서 Excel Wirte MAIN Methods ----

        public int ExcelWrite2(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste2(mPrinting, 1);
                    vPageRowCount = mCurrentRow2 - 1;    //첫장에 대해서는 시작row부터 체크.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //라인 열수.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite2(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
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

        #region ----- Foreign Currency - Excel Wirte MAIN Methods ----

        //public int ExcelWrite2(object pBalance_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pPayment)
        //{// 실제 호출되는 부분.

        //    string vMessage = string.Empty;

        //    int vTotalRow = 0;
        //    int vPageRowCount = 0;
        //    int vLIneRow = 0;
        //    bool vPrint_Flag = false;
        //    try
        //    {
        //        // 실제인쇄되는 행수.
        //        vTotalRow = pPayment.OraSelectData.Rows.Count;

        //        //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
        //        //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
        //        // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

        //        #region ----- Line Write ----

        //        if (vTotalRow > 0)
        //        {
        //            HeaderWrite1(pBalance_Date);
        //            // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
        //            mCopyLineSUM = CopyAndPaste2(mPrinting, 1);
        //            vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

        //            mCurrentRow = 6;
        //            mPrinting.XLActiveSheet(mTargetSheet);
        //            //SetArray1(pGrid, out vGDColumn, out vXLColumn);
        //            foreach (System.Data.DataRow vRow in pPayment.OraSelectData.Rows)
        //            {
        //                vLIneRow++;
        //                vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
        //                mAppInterface.OnAppMessageEvent(vMessage);
        //                System.Windows.Forms.Application.DoEvents();

        //                //계정코드 동일 여부 체크.
        //                vPrint_Flag = true;
        //                if (mAccount_Code == null || mAccount_Code == string.Empty || mIsNewPage == true)
        //                {
        //                    mMerger_Start = mCurrentRow;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                else if (mAccount_Code != iString.ISNull(vRow["ACCOUNT_CODE"]))
        //                {

        //                    mPrinting.XLCellMerge(mMerger_Start, 1, mMerger_End, 4, true);
        //                    mMerger_Start = mCurrentRow;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                else
        //                {
        //                    vPrint_Flag = false;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                mAccount_Code = iString.ISNull(vRow["ACCOUNT_CODE"]);

        //                mCurrentRow = LineWrite2(vRow, mCurrentRow, vPrint_Flag); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
        //                vPageRowCount = vPageRowCount + 1;

        //                if (vLIneRow == vTotalRow)
        //                {
        //                    // 마지막 데이터 이면 처리할 사항 기술
        //                    // 라인지운다 또는 합계를 표시한다 등 기술.
        //                    //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //합계.
        //                }
        //                else
        //                {
        //                    IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
        //                    if (mIsNewPage == true)
        //                    {
        //                        mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
        //                        vPageRowCount = mDefaultPageRow;
        //                    }
        //                }
        //            }
        //        }

        //        #endregion;
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mPrinting.XLOpenFileClose();
        //        mPrinting.XLClose();
        //    }
        //    return mPageNumber;
        //}

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                iDefaultEndRow = mCopy_EndRow - (mPrintingLastRow + 2);
                mCopyLineSUM = CopyAndPaste(mPrinting, mCurrentRow + iDefaultEndRow);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;
            string vActiveSheet = mSourceSheet1;

            mPageNumber = mPageNumber + 1;
            //if (mPageNumber > 1)
            //{
            //    2번째 인쇄페이지가 다른 양식일 경우 사용.
            //    vActiveSheet = mSourceSheet2;   
            //}

            // page수 표시.
            //XLPageNumber(pActiveSheet, mPageNumber);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

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

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste_SUB(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;
            string vActiveSheet = mSourceSheet2;

            mPageNumber = mPageNumber + 2;
            //if (mPageNumber > 1)
            //{
            //    2번째 인쇄페이지가 다른 양식일 경우 사용.
            //    vActiveSheet = mSourceSheet2;   
            //}

            // page수 표시.
            //XLPageNumber(pActiveSheet, mPageNumber);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(59, 1, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

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

        private int CopyAndPaste2(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow2;
            string vActiveSheet = mSourceSheet1;

            mPageNumber = mPageNumber + 1;
            
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

            return vPasteEndRow;
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        //private int CopyAndPaste2(XL.XLPrint pPrinting, int pPasteStartRow)
        //{
        //    int vPasteEndRow = pPasteStartRow + mCopy_EndRow2;
        //    string vActiveSheet = mSourceSheet1;

        //    mPageNumber = mPageNumber + 1;
        //    //if (mPageNumber > 1)
        //    //{
        //    //    2번째 인쇄페이지가 다른 양식일 경우 사용.
        //    //    vActiveSheet = mSourceSheet2;   
        //    //}

        //    // page수 표시.
        //    //XLPageNumber(pActiveSheet, mPageNumber);

        //    //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
        //    //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    pPrinting.XLActiveSheet(vActiveSheet);
        //    object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);

        //    //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
        //    //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    pPrinting.XLActiveSheet(mTargetSheet);
        //    object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol2, vPasteEndRow, mCopy_EndCol2);
        //    pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

        //    return vPasteEndRow;


        //    //int vCopySumPrintingLine = pCopySumPrintingLine;

        //    //int vCopyPrintingRowSTART = vCopySumPrintingLine;
        //    //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
        //    //int vCopyPrintingRowEnd = vCopySumPrintingLine;

        //    //pPrinting.XLActiveSheet("SourceTab1");
        //    //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    //pPrinting.XLActiveSheet("Destination");
        //    //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
        //    //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


        //    //mPageNumber++; //페이지 번호
        //    //// 페이지 번호 표시.
        //    ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
        //    ////int vRowSTART = vCopyPrintingRowEnd - 2;
        //    ////int vRowEND = vCopyPrintingRowEnd - 2;
        //    ////int vColumnSTART = 30;
        //    ////int vColumnEND = 33;
        //    ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
        //    ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //페이지 번호, XLcell[행, 열]

        //    //return vCopySumPrintingLine;
        //}

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
            mPrinting.XLPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            if (iString.ISNull(pSaveFileName) == string.Empty)
            {
                return;
            }

            //int vMaxNumber = MaxIncrement(pSavePath.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", pSavePath, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
            mPrinting.XLSave(pSaveFileName);

            //전호수 주석 처리 : 저장 방법 변경.
            //if (pSaveFileName == string.Empty)
            //{
            //    return;
            //}
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = 1; //MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
