using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0789
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
        private string mSourceSheet1 = "SourceTab1";
        private string mSourceSheet2 = "SourceTab2";

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

        public void HeaderWrite(System.Data.DataRow pRow, InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_SLC_DOC_ITEM_A)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;
            object vValue = null;
            string vString = string.Empty;

            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                ///////////
                vXLine = 7;

                //신고구분-정기
                vXLColumn = 10;
                vString = null;
                vValue = pRow["SLC_DOC_TYPE_01"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                } 
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //신고구분-수정
                vXLColumn = 15;
                vString = null;
                vValue = pRow["SLC_DOC_TYPE_02"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                } 
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //지급연월 
                vXLColumn = 22;
                vString = null;
                vValue = pRow["PAY_YYYYMM"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                } 
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //일괄납부-여 
                vXLColumn = 37;
                vString = null;
                vValue = pRow["PAYMENT_ALL_Y"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                } 
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);
                 
                //일괄납부-부
                vXLColumn = 41;
                vString = null;
                vValue = pRow["PAYMENT_ALL_N"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                } 
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 11;

                //법인명
                vXLColumn = 9;
                vString = null;
                vValue = pRow["CORP_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //대표자명
                vXLColumn = 28;
                vString = null;
                vValue = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 13;

                //사업자등록번호
                vXLColumn = 9;
                vString = null;
                vValue = pRow["VAT_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소재지
                vXLColumn = 28;
                vString = null;
                vValue = pRow["ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 15;

                //전화번호
                vXLColumn = 9;
                vString = null;
                vValue = pRow["TEL_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //전자우편
                vXLColumn = 28;
                vString = null;
                vValue = pRow["EMAIL"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //-- 원천공제 명세 및 납부할 금액 --//
                vXLine = 21;
                foreach (System.Data.DataRow vROW in pIDA_SLC_DOC_ITEM_A.CurrentRows)
                {
                    //소득구분
                    vXLColumn = 1;
                    vString = null;
                    vValue = vROW["SLC_INCOME_TYPE_NAME"];
                    if (iString.ISNull(vValue) != string.Empty)
                    {
                        vString = string.Format("{0}", vValue);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    //소득구분코드
                    vXLColumn = 7;
                    vString = null;
                    vValue = vROW["SLC_INCOME_TYPE"];
                    if (iString.ISNull(vValue) != string.Empty)
                    {
                        vString = string.Format("{0}", vValue);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    //원촌공제통지인원
                    vXLColumn = 11;
                    vString = null;
                    vValue = vROW["ORI_SLC_PERSON_COUNT"];
                    if (iString.ISDecimal(vValue) == true)
                    {
                        vString = string.Format("{0:###,###,###,###,###,###,###}", vValue); 
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    //원천공제통지금액
                    vXLColumn = 18;
                    vString = null;
                    vValue = vROW["ORI_SLC_AMOUNT"];
                    if (iString.ISDecimal(vValue) == true)
                    {
                        vString = string.Format("{0:###,###,###,###,###,###,###}", vValue);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    //원천공제인원
                    vXLColumn = 25;
                    vString = null;
                    vValue = vROW["PAY_PERSON_COUNT"];
                    if (iString.ISDecimal(vValue) == true)
                    {
                        vString = string.Format("{0:###,###,###,###,###,###,###}", vValue);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    //원천공제금액
                    vXLColumn = 31;
                    vString = null;
                    vValue = vROW["PAY_SLC_AMOUNT"];
                    if (iString.ISDecimal(vValue) == true)
                    {
                        vString = string.Format("{0:###,###,###,###,###,###,###}", vValue);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                    vXLine = vXLine + 2;
                } 

                //--------------------------------------------//
                vXLine = 35;

                //일자
                vXLColumn = 31;
                vString = null;
                vValue = pRow["SUBMIT_REPORT_DATE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 37;

                //원천공제의무자
                vXLColumn = 25;
                vString = null;
                vValue = pRow["SUBMIT_REPORTER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 39;

                //귀하
                vXLColumn = 1;
                vString = null;
                vValue = pRow["TAX_OFFIECER_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //--------------------------------------------//
                vXLine = 44;

                //환급금융기관
                vXLColumn = 7;
                vString = null;
                vValue = pRow["REFUND_BANK_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //예금종류
                vXLColumn = 20;
                vString = null;
                vValue = pRow["REFUND_DEPOSIT_TYPE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //계좌번호
                vXLColumn = 32;
                vString = null;
                vValue = pRow["REFUND_ACCOUNT_NUM"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //상환금명세서(을) 헤더.
                mPrinting.XLActiveSheet(mSourceSheet2);

                //--------------------------------------------//
                vXLine = 7;

                //법인명
                vXLColumn = 10;
                vString = null;
                vValue = pRow["CORP_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //사업자등록번호
                vXLColumn = 31;
                vString = null;
                vValue = pRow["VAT_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
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
            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                //순번
                vConvertString = null;
                vObject = pRow["SEQ"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject); 
                } 
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //성명
                vConvertString = null;
                vObject = pRow["NAME"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                vXLColumn = 3;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //주민번호
                vConvertString = null;
                vObject = pRow["REPRE_NUM"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득구분
                vConvertString = null;
                vObject = pRow["SLC_INCOME_TYPE_NAME"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득구분코드
                vConvertString = null;
                vObject = pRow["SLC_INCOME_TYPE"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원천공제통지액
                vConvertString = null;
                vObject = pRow["ORI_SLC_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                } 
                vXLColumn = 20;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전월미공제액
                vConvertString = null;
                vObject = pRow["PRE_PAY_SLC_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                } 
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원천공제액
                vConvertString = null;
                vObject = pRow["PAY_SLC_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                } 
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //이월공제액
                vConvertString = null;
                vObject = pRow["NEXT_PAY_SLC_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                } 
                vXLColumn = 32;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //사유코드
                vConvertString = null;
                vObject = pRow["SLC_REASON_CODE"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원천공제일
                vConvertString = null;
                vObject = pRow["PAY_SUPPLY_DATE"];
                if (iString.ISNull(vObject) != String.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }  
                vXLColumn = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString); 

                //-------------------------------------------------------------------
                vXLine++;
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

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_SLC_DOC
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_SLC_DOC_ITEM_A
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_SLC_DOC_ITEM_B)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            mPageNumber = 0;
            mCopyLineSUM = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 43;
            mCopy_EndRow = 59;

            mCopy_StartCol2 = 1;
            mCopy_StartRow2 = 1;
            mCopy_EndCol2 = 43;
            mCopy_EndRow2 = 59;
            mPrintingLastRow = 57;  //최종 인쇄 라인.

            mCurrentRow = 1;
            mDefaultPageRow2 = 14;  //2번째장.

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
             
            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pIDA_SLC_DOC.CurrentRows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1); 

                    mPrinting.XLActiveSheet(mTargetSheet);
                    HeaderWrite(pIDA_SLC_DOC.CurrentRow, pIDA_SLC_DOC_ITEM_A);

                    mCurrentRow = mCopy_EndRow + 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow);
                    mCurrentRow = mCurrentRow + mDefaultPageRow2;

                    vPageRowCount = mDefaultPageRow2 + 1;
                    foreach (System.Data.DataRow vRow in pIDA_SLC_DOC_ITEM_B.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        
                        mCurrentRow = LineWrite(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 2;

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
                                vPageRowCount = mDefaultPageRow + 1;
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
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pSourceTab, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(pSourceTab); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab);
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

        private int CopyAndPaste2(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow2;
            string vActiveSheet = mSourceSheet2;

            mPageNumber = mPageNumber + 1;
            
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(mCurrentRow, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);
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
