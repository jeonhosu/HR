﻿using System;
using ISCommonUtil;

namespace HRMF0506
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

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 63;
        private int mCopy_EndRow = 34;
        private int mPrintingLastRow = 33;  //실제 데이터 인쇄 최종 라인.

        private int mCurrentRow = 7;        //실제 인쇄되는 row 위치.
        private int mDefaultPageRow = 6;    //페이지 skip후 적용되는 기본 PageCount 기본값.

        //총합계 : 건수, 공급가액, 세액.
        private decimal mTOT_COUNT = 0;
        private decimal mTOT_GL_AMOUNT = 0;
        private decimal mTOT_VAT_AMOUNT = 0;

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

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[13];
            pXLColumn = new int[13];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("PAY_YYYYMM");
            pGDColumn[1] = pGrid.GetColumnToIndex("PAY_AMOUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("BONUS_AMOUNT");
            pGDColumn[3] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("PENSION_INSUR");
            pGDColumn[5] = pGrid.GetColumnToIndex("HEALTH_INSUR");
            pGDColumn[6] = pGrid.GetColumnToIndex("EMPLOYMENT_INSUR");
            pGDColumn[7] = pGrid.GetColumnToIndex("INCOME_TAX");
            pGDColumn[8] = pGrid.GetColumnToIndex("RESIDENTS_TAX");
            pGDColumn[9] = pGrid.GetColumnToIndex("DED_PAY_AMOUNT");
            pGDColumn[10] = pGrid.GetColumnToIndex("DED_BONUS_AMOUNT");
            pGDColumn[11] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");
            pGDColumn[12] = pGrid.GetColumnToIndex("REAL_AMOUNT");


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 1;
            pXLColumn[1] = 4;
            pXLColumn[2] = 9;
            pXLColumn[3] = 14;
            pXLColumn[4] = 19;
            pXLColumn[5] = 24;
            pXLColumn[6] = 29;
            pXLColumn[7] = 34;
            pXLColumn[8] = 39;
            pXLColumn[9] = 44;
            pXLColumn[10] = 49;
            pXLColumn[11] = 54;
            pXLColumn[12] = 59;
        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[13]; 
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("PAY_YYYYMM");
            pGDColumn[1] = pGrid.GetColumnToIndex("PAY_AMOUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("BONUS_AMOUNT");
            pGDColumn[3] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("PENSION_INSUR");
            pGDColumn[5] = pGrid.GetColumnToIndex("HEALTH_INSUR");
            pGDColumn[6] = pGrid.GetColumnToIndex("EMPLOYMENT_INSUR");
            pGDColumn[7] = pGrid.GetColumnToIndex("INCOME_TAX");
            pGDColumn[8] = pGrid.GetColumnToIndex("RESIDENTS_TAX");
            pGDColumn[9] = pGrid.GetColumnToIndex("DED_PAY_AMOUNT");
            pGDColumn[10] = pGrid.GetColumnToIndex("DED_BONUS_AMOUNT");
            pGDColumn[11] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");
            pGDColumn[12] = pGrid.GetColumnToIndex("REAL_AMOUNT"); 
        }

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

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataCommand pPrinted_Value, object vDEPT_NAME, object vPAY_GRADE_NAME, object vPERSON_NAME)
        {// 헤더 인쇄.
            object vPrinted_Value;
            object vPrinted_Value2;
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // title
                //vXLine = 1;
                //vXLColumn = 1;
                //mPrinting.XLSetCell(vXLine, vXLColumn, pTitle);

                //corporation
                vXLine = 3;
                vXLColumn = 1;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("O_CORP_NAME");
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //period
                vXLine = 3;
                vXLColumn = 33;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("P_DATE_FR");
                vPrinted_Value2 = pPrinted_Value.GetCommandParamValue("P_DATE_TO");

                vPrinted_Value = string.Format("기간 : {0} ~ {1}", vPrinted_Value, vPrinted_Value2);

                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //period
                vXLine = 4;
                vXLColumn = 1;

                vPrinted_Value = string.Format("{0} {1} {2}", vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);

                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //printed date
                vXLine = 36;
                vXLColumn = 1;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("O_PRINTED_DATE");
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_2(InfoSummit.Win.ControlAdv.ISDataCommand pPrinted_Value, object vDEPT_NAME, object vPAY_GRADE_NAME, object vPERSON_NAME)
        {// 헤더 인쇄.
            object vPrinted_Value;
            object vPrinted_Value2;
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                //corporation
                vXLine = 3;
                vXLColumn = 1;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("O_CORP_NAME");
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //period
                vXLine = 3;
                vXLColumn = 23;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("P_DATE_FR");
                vPrinted_Value2 = pPrinted_Value.GetCommandParamValue("P_DATE_TO");

                vPrinted_Value = string.Format("기간 : {0} ~ {1}", vPrinted_Value, vPrinted_Value2);

                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //period
                vXLine = 4;
                vXLColumn = 1;

                vPrinted_Value = string.Format("{0}/{1}/{2}", vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);

                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);

                //printed date
                vXLine = 28;
                vXLColumn = 1;
                vPrinted_Value = pPrinted_Value.GetCommandParamValue("O_PRINTED_DATE");
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrinted_Value);
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

        #region ----- Excel Write [Line] Method -----

        private int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;

            //숫자 포맷 적용 예.
            //decimal vConvertDecimal = 0m;
            //DateTime vCONVERT_DATE = new DateTime(); ;
            //vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //0 - 년월
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[0]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[0], vConvertString);

                //1 - 급여액
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[1]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[1], vConvertString);

                //2-상여액
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[2]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[2], vConvertString);

                //3-총금액.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[3]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[3], vConvertString);


                //4-국민연금.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[4]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[4], vConvertString);

                //5-건강보험
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[5]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[5], vConvertString);

                //6-고용보험.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[6]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[6], vConvertString);

                //7-소득세.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[7]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[7], vConvertString);


                //8-주민세.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[8]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[8], vConvertString);


                //9-급여공제.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[9]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[9], vConvertString);

                //10-상여공제.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[10]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[10], vConvertString);

                //11-총공제
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[11]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[11], vConvertString);

                //12-실지급액.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[12]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, pXLColumn[12], vConvertString);

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

        private int LineWrite_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;

            //숫자 포맷 적용 예.
            //decimal vConvertDecimal = 0m;
            //DateTime vCONVERT_DATE = new DateTime(); ;
            //vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //0 - 년월
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[0]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //1 - 급여액
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[1]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 5, vConvertString);

                //2-상여액
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[2]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vConvertString);

                //3-총금액.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[3]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vConvertString);


                //4-국민연금.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[4]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 24, vConvertString);

                //5-건강보험
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[5]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //6-고용보험.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[6]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vConvertString);

                //7-소득세.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[7]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 39, vConvertString);


                //8-주민세.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[8]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vConvertString);


                //9-급여공제.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[9]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 49, vConvertString);

                //10-상여공제.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[10]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 54, vConvertString);

                //11-총공제
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[11]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 59, vConvertString);

                //12-실지급액.
                vObject = pGrid.GetCellValue(pGridRow, pGDColumn[12]);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 64, vConvertString);

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

        #region ----- Excel Wirte MAIN Methods ----

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataCommand pPrinted_Value, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, object vDEPT_NAME, object vPAY_GRADE_NAME, object vPERSON_NAME)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int[] vGDColumn;
            int[] vXLColumn;
            int vTotalRow = 0;
            int vPageRowCount = 0;
            try
            {
                mCopy_StartCol = 1;
                mCopy_StartRow = 1;
                mCopy_EndCol = 70;
                mCopy_EndRow = 29;
                mPrintingLastRow = 28;  //실제 데이터 인쇄 최종 라인.

                mCurrentRow = 6;        //실제 인쇄되는 row 위치.
                mDefaultPageRow = 5;    //페이지 skip후 적용되는 기본 PageCount 기본값. 

                HeaderWrite(pPrinted_Value, vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);
                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                // 총합계.
                mTOT_COUNT = 0;
                mTOT_GL_AMOUNT = 0;
                mTOT_VAT_AMOUNT = 0;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite(pGrid, vRow, mCurrentRow, vGDColumn, vXLColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(9);      //합계.
                            //mCurrentRow = XLTOTAL_Line(13);     // 수출재화 합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mCopy_EndRow - mPrintingLastRow + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
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
            if (mPageNumber == 0)
            {
                mPageNumber = 1;
            }
            return mPageNumber;
        }

        public int ExcelWrite_2(InfoSummit.Win.ControlAdv.ISDataCommand pPrinted_Value, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, object vDEPT_NAME, object vPAY_GRADE_NAME, object vPERSON_NAME)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            int[] vGDColumn;
            int[] vXLColumn;
            int vTotalRow = 0;
            int vPageRowCount = 0;
            try
            {
                mCopy_StartCol = 1;
                mCopy_StartRow = 1;
                mCopy_EndCol = 70;
                mCopy_EndRow = 29;
                mPrintingLastRow = 28;  //실제 데이터 인쇄 최종 라인.

                mCurrentRow = 6;        //실제 인쇄되는 row 위치.
                mDefaultPageRow = 5;    //페이지 skip후 적용되는 기본 PageCount 기본값. 

                HeaderWrite_2(pPrinted_Value, vDEPT_NAME, vPAY_GRADE_NAME, vPERSON_NAME);
                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                // 총합계.
                mTOT_COUNT = 0;
                mTOT_GL_AMOUNT = 0;
                mTOT_VAT_AMOUNT = 0;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray2(pGrid, out vGDColumn);
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_2(pGrid, vRow, mCurrentRow, vGDColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            //mCurrentRow = XLTOTAL_Line(9);      //합계.
                            //mCurrentRow = XLTOTAL_Line(13);     // 수출재화 합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mCopy_EndRow - mPrintingLastRow + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
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
            if (mPageNumber == 0)
            {
                mPageNumber = 1;
            }
            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                iDefaultEndRow = mCopy_EndRow - mPrintingLastRow;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow + iDefaultEndRow);
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

            // page수 표시.
            mPageNumber = mPageNumber + 1;
            XLPageNumber(pActiveSheet, mPageNumber);

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);
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

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

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

        #endregion;
    }
}
