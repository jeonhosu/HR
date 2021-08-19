using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0775
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
        private int mCopy_EndCol = 41;
        private int mCopy_EndRow = 45;
        private int mPrintingLastRow = 35;  //실제 데이터 인쇄 최종 라인.

        private int mCurrentRow = 20;        //실제 인쇄되는 row 위치.
        private int mDefaultPageRow = 19;    //페이지 skip후 적용되는 기본 PageCount 기본값.

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

                //소득자 보관용 여부 
                vXLine = 6;
                vXLColumn = 10;
                vValue = pRow["PRINT_TYPE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                    if (vString == "1")
                    {
                        vString = "[√]";
                    }
                    else
                    {
                        vString = "[ ]";
                    }
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //발행자 보관용 여부 
                vXLine = 6;
                vXLColumn = 18;
                vValue = pRow["PRINT_TYPE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                    if (vString == "2")
                    {
                        vString = "[√]";
                    }
                    else
                    {
                        vString = "[ ]";
                    }
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //내외국인-내국인
                vXLine = 5;
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
                vXLine = 6;
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

                //소득자 성명
                vXLine = 10;
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
                vXLine = 10;
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
                vXLine = 11;
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

                //소득구분 68 
                vXLine = 12;
                vXLColumn = 7;
                vValue = pRow["INCOME_68"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_69
                vXLine = 12;
                vXLColumn = 16;
                vValue = pRow["INCOME_69"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_63
                vXLine = 12;
                vXLColumn = 25;
                vValue = pRow["INCOME_63"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_60
                vXLine = 13;
                vXLColumn = 7;
                vValue = pRow["INCOME_60"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_64
                vXLine = 13;
                vXLColumn = 21;
                vValue = pRow["INCOME_64"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_71
                vXLine = 14;
                vXLColumn = 7;
                vValue = pRow["INCOME_71"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_72
                vXLine = 14;
                vXLColumn = 14;
                vValue = pRow["INCOME_72"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_73
                vXLine = 14;
                vXLColumn = 21;
                vValue = pRow["INCOME_73"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_74
                vXLine = 14;
                vXLColumn = 27;
                vValue = pRow["INCOME_74"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_75
                vXLine = 15;
                vXLColumn = 7;
                vValue = pRow["INCOME_75"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_76
                vXLine = 15;
                vXLColumn = 14;
                vValue = pRow["INCOME_76"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //소득구분_62
                vXLine = 16;
                vXLColumn = 7;
                vValue = pRow["INCOME_62"];
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
                vXLine = 37;
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
                vXLine = 38;
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
                vXLine = 39;
                vXLColumn = 1;
                vValue = pRow["RECEIVER_NAME"];
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
                //지급연
                vXLColumn = 1;
                vObject = pRow["PAY_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //지급월
                vXLColumn = 3;
                vObject = pRow["PAY_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지급일
                vXLColumn = 5;
                vObject = pRow["PAY_DAY"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득귀속연
                vXLColumn = 7;
                vObject = pRow["RECEIPT_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득귀속월.
                vXLColumn = 9;
                vObject = pRow["RECEIPT_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //지급총액.
                vXLColumn = 11;
                vObject = pRow["TOT_PAYMENT_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //필요경비.
                vXLColumn = 16;
                vObject = pRow["EXP_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득금액.
                vXLColumn = 20;
                vObject = pRow["TOT_INCOME_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //세율.
                vXLColumn = 24;
                vObject = pRow["TAX_RATE"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //소득세
                vXLColumn = 26;
                vObject = pRow["INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //지방소득세
                vXLColumn = 30;
                vObject = pRow["LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //농특세
                vXLColumn = 34;
                vObject = pRow["SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계
                vXLColumn = 38;
                vObject = pRow["TOTAL_DED_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
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

        #endregion;

        #region ----- Excel Write [CURRENCY] Method -----

        private int LineWrite2(System.Data.DataRow pRow, int pXLine, bool pPrint_Flag)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
                if (pPrint_Flag == true)
                {
                    //계정명
                    vXLColumn = 1;
                    vObject = pRow["ACCOUNT_DESC"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vConvertString = string.Format("{0}", vObject);

                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                }
                //거래처코드
                vXLColumn = 5;
                vObject = pRow["CUSTOMER_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //거래처 명.
                vXLColumn = 9;
                vObject = pRow["CUSTOMER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //잔액월.
                vXLColumn = 17;
                vObject = pRow["BALANCE_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //통화.
                vXLColumn = 19;
                vObject = pRow["CURRENCY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외화금액.
                vXLColumn = 22;
                vObject = pRow["GL_CURR_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###.##}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원화금액.
                vXLColumn = 28;
                vObject = pRow["GL_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //현금외화지급액.
                vXLColumn = 33;
                vObject = pRow["CASH_CURR_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###.##}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //현금지급액.
                vXLColumn = 39;
                vObject = pRow["CASH_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
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

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader, InfoSummit.Win.ControlAdv.ISDataAdapter pLine)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
            //bool vPrint_Flag = false;
            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pHeader.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    //헤더 인쇄.
                    foreach (System.Data.DataRow vRow in pHeader.OraSelectData.Rows)
                    {
                        HeaderWrite(vRow);
                    }
                    
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                    vTotalRow = pLine.OraSelectData.Rows.Count;  //라인 열수.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pLine.OraSelectData.Rows)
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
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
            //mPrinting.XLPrinting(pPageSTART, pPageEND, 1);
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
