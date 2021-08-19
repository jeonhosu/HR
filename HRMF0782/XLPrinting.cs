using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0782
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
        private int mCopy_EndCol = 33;
        private int mCopy_EndRow = 27;
        private int mPrintingLastRow = 27;  //실제 데이터 인쇄 최종 라인.

        private int mCurrentRow = 12;        //실제 인쇄되는 row 위치.
        private int mDefaultPageRow = 11;    //페이지 skip후 적용되는 기본 PageCount 기본값.

        // 인쇄2 - 소득세 납부서 인쇄 정보.
        private int mCopy_StartCol2 = 1;
        private int mCopy_StartRow2 = 1;
        private int mCopy_EndCol2 = 68;
        private int mCopy_EndRow2 = 38;
        private int mPrintingLastRow2 = 38;  //실제 데이터 인쇄 최종 라인.

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
                //귀속월
                vObject = pRow["STD_MM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 4;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지급일
                vObject = pRow["PAY_SUPPLY_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 4;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 인원수
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 과세표준
                vObject = pRow["A01_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 지방소득세
                vObject = pRow["A01_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 인원수
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 과세표준
                vObject = pRow["A02_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 지방소득세
                vObject = pRow["A02_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 인원수
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 과세표준
                vObject = pRow["A03_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 지방소득세
                vObject = pRow["A03_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득 인원수
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득 과세표준
                vObject = pRow["A04_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득 지방소득세
                vObject = pRow["A04_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 인원수
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 과세표준
                vObject = pRow["A05_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 지방소득세
                vObject = pRow["A05_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 인원수
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 과세표준
                vObject = pRow["A06_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 지방소득세
                vObject = pRow["A06_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 인원수
                vObject = pRow["A07_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 과세표준
                vObject = pRow["A07_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 지방소득세
                vObject = pRow["A07_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터받은소득 인원수
                vObject = pRow["A08_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터받은소득 과세표준
                vObject = pRow["A08_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터받은소득 지방소득세
                vObject = pRow["A08_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 인원수
                vObject = pRow["A09_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 과세표준
                vObject = pRow["A09_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 지방소득세
                vObject = pRow["A09_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 인원수
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 과세표준
                vObject = pRow["A10_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 지방소득세
                vObject = pRow["A10_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //가감세액(조정액)지방소득세
                vObject = pRow["TOTAL_ADJUST_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 인원수
                vObject = pRow["A90_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 과세표준
                vObject = pRow["A90_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 지방소득세
                vObject = pRow["PAY_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //제출일자
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업장 주소
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //상호
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자등록번호
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //대표자
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전화번호
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 24;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신고하는 시군구
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 13;
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

        #region ----- Excel Write [CURRENCY] Method -----

        private int LineWrite2(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
                //기관코드                
                vObject = pRow["TAX_OFFICE_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 4;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //회계코드
                vObject = pRow["TAX_ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 4;
                vXLColumn = 16;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 62;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //과세년월
                vObject = pRow["TAX_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 5;
                vXLColumn = 5;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부기한.
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 5;
                vXLColumn = 16;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 62;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //상호
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 6;
                vXLColumn = 11;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57; 
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //주민(법인)등록번호
                vObject = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 7;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //대표자
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 8;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업자등록번호
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 9;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //주소
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 10;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전화번호
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 11;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //귀속년월.
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 12;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 47;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신고하는 시군구.
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 13;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //납부액
                vObject = pRow["PAY_LOCAL_TAX_KOR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 14;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 53;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 인원수
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 과세표준
                vObject = pRow["A01_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이자소득 지방소득세
                vObject = pRow["A01_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 인원수
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 과세표준
                vObject = pRow["A02_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //배당소득 지방소득세
                vObject = pRow["A02_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 인원수
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 과세표준
                vObject = pRow["A03_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사업소득 지방소득세
                vObject = pRow["A03_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //근로소득 인원수
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득 과세표준
                vObject = pRow["A04_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //근로소득 지방소득세
                vObject = pRow["A04_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 인원수
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 과세표준
                vObject = pRow["A05_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //연금소득 지방소득세
                vObject = pRow["A05_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 인원수
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 과세표준
                vObject = pRow["A06_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기타소득 지방소득세
                vObject = pRow["A06_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 인원수
                vObject = pRow["A07_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 과세표준
                vObject = pRow["A07_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //퇴직소득 지방소득세
                vObject = pRow["A07_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터 받은소득 인원수
                vObject = pRow["A08_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터 받은소득 과세표준
                vObject = pRow["A08_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외국인으로부터 받은소득 지방소득세
                vObject = pRow["A08_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 인원수
                vObject = pRow["A09_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 과세표준
                vObject = pRow["A09_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //법인세법 제98조 지방소득세
                vObject = pRow["A09_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 인원수
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 과세표준
                vObject = pRow["A10_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세법 제119조 지방소득세
                vObject = pRow["A10_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //가감세액(조정액)
                vObject = pRow["TOTAL_ADJUST_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 29;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 인원수
                vObject = pRow["A90_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 과세표준
                vObject = pRow["A90_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계 지방소득세
                vObject = pRow["PAY_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //제출일자.
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 32;
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                vXLine = 33;
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //서명 또는 인.
                vObject = pRow["REPORT_CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 33;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //귀하
                vObject = pRow["TAX_OFFIECER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 37;
                vXLColumn = 5;
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
