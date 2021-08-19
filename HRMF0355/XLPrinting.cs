using System;
using ISCommonUtil;

namespace HRMF0355
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
        private string mTargetSheet = "PRINTING";
        private string mSourceSheet1 = "SOURCE1";
        private string mSourceSheet2 = "SOURCE2";

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
        private int mCopy_EndCol = 43;
        private int mCopy_EndRow = 64;
        private int mPrintingLastRow = 62;  //최종 인쇄 라인.

        private int mCurrentRow = 5;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 4;    //페이지 skip후 적용되는 기본 PageCount 기본값.

        //본부코드, 작업장코드
        private string gDEPT_GROUP_CODE = null;
        private string gFLOOR_CODE = null;

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
            pGDColumn = new int[10];
            pXLColumn = new int[10];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = 0;            
            pGDColumn[1] = pGrid.GetColumnToIndex("DEPT_2ND_DESC");
            pGDColumn[2] = pGrid.GetColumnToIndex("FLOOR_NAME");
            pGDColumn[3] = pGrid.GetColumnToIndex("NAME");
            pGDColumn[4] = pGrid.GetColumnToIndex("POST_NAME");
            pGDColumn[5] = pGrid.GetColumnToIndex("IN_TIME");            
            pGDColumn[6] = pGrid.GetColumnToIndex("DUTY_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("DESCRIPTION");
            pGDColumn[8] = pGrid.GetColumnToIndex("LATE_DESC");
            pGDColumn[9] = pGrid.GetColumnToIndex("OUT_TIME");


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 1;
            pXLColumn[1] = 3;
            pXLColumn[2] = 10;
            pXLColumn[3] = 16;
            pXLColumn[4] = 21;
            pXLColumn[5] = 25;
            pXLColumn[6] = 30;
            pXLColumn[7] = 30;
            pXLColumn[8] = 30;
            pXLColumn[9] = 39;
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

        #region ----- Excel Write : Header Write Method ----

        public void HeaderWrite(object pWORK_DATE, object pWEEK_DESC)
        {// 헤더 인쇄.
            object vWORK_DATE = String.Format("({0} {1})", pWORK_DATE, pWEEK_DESC);
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);
                // 근무일자.
                vXLine = 1;
                vXLColumn = 4;
                mPrinting.XLSetCell(vXLine, vXLColumn, vWORK_DATE);
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

        private void XLHeader1(object pWORK_DATE, object pWEEK_DESC)
        {// 헤더 인쇄.
            object vWORK_DATE = String.Format("({0} {1})", pWORK_DATE, pWEEK_DESC);
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);
                // 근무일자.
                vXLine = 1;
                vXLColumn = 4;
                mPrinting.XLSetCell(vXLine, vXLColumn, vWORK_DATE);
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

        private int XLLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, 
                            int pIDX_DEPT_GROUP_CODE, int pIDX_FLOOR_CODE, Boolean pNewPage)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            object vTEMP_CODE = null;
            object vDESCRIPTION = null;
            object vLATE_DESC = null;

            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //DateTime vCONVERT_DATE = new DateTime(); ;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //0 - 일련번호
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = Convert.ToDecimal(pGridRow) + 1;
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //1 - 본부명
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                if (pNewPage == true || pGridRow == 0 || gDEPT_GROUP_CODE != iString.ISNull(pGrid.GetCellValue(pGridRow, pIDX_DEPT_GROUP_CODE)))
                {
                    
                }
                else
                {
                    vObject = null;
                    mPrinting.XL_LineClearTOP(mCurrentRow, 3, 9);
                }
                vTEMP_CODE = vObject;
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //2-작업장명
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                if (iString.ISNull(vTEMP_CODE) != string.Empty || gFLOOR_CODE != iString.ISNull(pGrid.GetCellValue(pGridRow, pIDX_FLOOR_CODE)))
                {
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                }
                else
                {
                    vObject = null;
                    mPrinting.XL_LineClearTOP(mCurrentRow, 10, 15);
                }
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //3-성명
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //4-직급
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //5-출근
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //6-판정
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);         //근태계명
                vGDColumnIndex = pGDColumn[7];
                vDESCRIPTION = pGrid.GetCellValue(pGridRow, vGDColumnIndex);    //비고
                vGDColumnIndex = pGDColumn[8];
                vLATE_DESC = pGrid.GetCellValue(pGridRow, vGDColumnIndex);      // 지각/조퇴 비고.
                if (iString.ISNull(vDESCRIPTION) != String.Empty)
                {
                    vObject = String.Format("{0} : {1}", vObject, vDESCRIPTION);
                }
                else if(iString.ISNull(vLATE_DESC) != String.Empty)
                {
                    vObject = vLATE_DESC;
                }
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //8-퇴근
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                //-------------------------------------------------------------------
                gDEPT_GROUP_CODE = iString.ISNull(pGrid.GetCellValue(pGridRow, pIDX_DEPT_GROUP_CODE));
                gFLOOR_CODE = iString.ISNull(pGrid.GetCellValue(pGridRow, pIDX_FLOOR_CODE));

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

        private int XLTOTAL_Line(int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            //int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            //int vXLColumnIndex = 0;

            //string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;
            //bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                ////12-건수
                //vXLColumnIndex = 12;
                //IsConvert = IsConvertNumber(mTOT_COUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////22-공급가액
                //vXLColumnIndex = 22;
                //IsConvert = IsConvertNumber(mTOT_GL_AMOUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////34-세액
                //vXLColumnIndex = 34;
                //IsConvert = IsConvertNumber(mTOT_VAT_AMOUNT, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                ////-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            //pXLine = vXLine;

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

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, object pWORK_DATE, object pWEEK_DESC)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            int[] vGDColumn;
            int[] vXLColumn;
            int vTotalRow = 0;
            int vPageRowCount = 0;
            try
            {                   
                XLHeader1(pWORK_DATE, pWEEK_DESC);  // 헤더 인쇄.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                // 실제인쇄되는 행수.    
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                // 본부코드, 작업장코드
                gDEPT_GROUP_CODE = null;
                gFLOOR_CODE = null;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               
                if (vTotalRow > 0)
                {
                    SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    int vIDX_DEPT_GROUP_CODE = pGrid.GetColumnToIndex("DEPT_2ND_CODE");
                    int vIDX_FLOOR_CODE = pGrid.GetColumnToIndex("FLOOR_CODE");

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        mCurrentRow = XLLine(pGrid, vRow, mCurrentRow, vGDColumn, vXLColumn, vIDX_DEPT_GROUP_CODE, vIDX_FLOOR_CODE, mIsNewPage); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCurrentRow + mPrintingLastRow - vPageRowCount, mCopy_EndCol);
                            mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, 2);
                            
                            //vPageRowCount = vPageRowCount + 1;
                            //for (int c = vPageRowCount; c < mPrintingLastRow; c++)
                            //{
                            //    if (c == vPageRowCount)
                            //    {
                            //        mPrinting.XL_LineDraw_Bottom(mCurrentRow - 1, mCopy_StartCol, mCopy_EndCol, 2);
                            //    }
                            //    mCurrentRow = mCurrentRow + 1;
                            //    mPrinting.XL_LineClearALL(mCurrentRow, mCopy_StartCol, mCurrentRow + mPrintingLastRow - vPageRowCount, mCopy_EndCol);
                            //}
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                //if (mPageNumber == 0)
                                //{
                                    mCurrentRow = mCurrentRow + mDefaultPageRow;
                                    vPageRowCount = mDefaultPageRow;
                                //}
                                //else
                                //{
                                //    mCurrentRow = mCurrentRow + 1;
                                //    vPageRowCount = 1;
                                //}
                            }
                        }
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
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

        #region ----- CellMerge -----

        private Boolean CellMerge(object pOLD_CODE, object pNEW_CODE, int pSTART_ROW, int pSTART_COL, int pEND_ROW, int pEND_COL)
        {
            //if(mIsNewPage == true)
            //{
            //    mPrinting.XLCellMerge(pSTART_ROW, pSTART_COL, pEND_ROW, pEND_COL, true);

            //    gDeptStartRow = mCurrentRow;
            //    gFloorStartRow = gDeptStartRow;
            //    return true;
            //}
            //else if (iString.ISNull(pOLD_CODE) != iString.ISNull(pNEW_CODE))
            //{
            //    mPrinting.XLCellMerge(pSTART_ROW, pSTART_COL, pEND_ROW, pEND_COL, true);
            //    return true;
            //}
            //else
            //{
            //    return false;
            //}
            return false;
        }

        #endregion

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                iDefaultEndRow = mCopy_EndRow - mPrintingLastRow;
                //if (mPageNumber == 0)
                //{
                //    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow + iDefaultEndRow);
                //}
                //else
                //{
                //    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                //}
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow + iDefaultEndRow);
                mCurrentRow = mCurrentRow + iDefaultEndRow;
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
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
