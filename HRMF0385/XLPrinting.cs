using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0385
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        // 쉬트명 정의.
        private string mTargetSheet = "Destination";
        private string mSourceSheet1 = "SourceTab1";
        private string mSourceSheet2 = "SourceTab2";

        private int mPageTotalNumber = 0;
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
        private int mPrintingLastRow = 0;  //최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultEndPageRow = 0;    // 페이지 증가후 PageCount 기본값.
        private int mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.

        private int mCountLinePrinting = 0; //엑셀 라인 Seq 

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

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
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

        #region ----- Convert DateTime Methods ----

        private object ConvertDateTime(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        private object ConvertDate(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
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

        #region ----- Content Clear All Methods ----

        private void XlAllContentClear(XL.XLPrint pPrinting)
        { 
            pPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = mPrintingLineSTART1;
            //int vStartCol = mCopyColumnSTART;
            //int vEndRow = mPrintingLineEND1 + 5;
            //int vEndCol = mCopyColumnEND - 1;

            //mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
        
        }

        #endregion;

        #region ----- Line Clear All Methods ----

        private void XlLineClear(int pPrintingLine)
        {
            mPrinting.XLActiveSheet("SourceTab1");

            //int vStartRow = pPrintingLine + 1;
            //int vStartCol = mCopyColumnSTART + 1;
            //int vEndRow = mPrintingLineEND1 - 4;
            //int vEndCol = mCopyColumnEND - 1;

            //if (vStartRow > vEndRow)
            //{
            //    vStartRow = vEndRow; //시작하는 행이 계산후, 끝나는 행 보다 값이 커지므로, 끝나는 행 값을 줌
            //}

            //mPrinting.XL_LineClearInSide(vStartRow, vStartCol, vEndRow, vEndCol);
            //mPrinting.XL_LineClearInSide(vEndRow + 2, vStartCol, vEndRow, vEndCol);
        
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(object pDUTY_YYYYMM, object pLOCAL_DATE, InfoSummit.Win.ControlAdv.ISGridAdvEx pData)
        {
            object vObject = null;
            string vString = string.Empty;
            int vLine = 4;

            // 쉬트명 정의.
            mTargetSheet = "Sheet1";
            mSourceSheet1 = "Source1";
            mSourceSheet2 = "Source1";

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1); //셀에 문자를 넣기 위해 쉬트 선택
                //작성부서명[DEPT_CODE DEPT_NAME]
                vObject = pDUTY_YYYYMM;
                if (vObject != null)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(1, 10, vString);
                 
                //인쇄일시[PRINTED DATE]
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vString = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(2, 33, vString);

                //IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].Visible = 1;
                //IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].Default = iString.ISNull(mCOLUMN_DESC);
                //IGR_MONTH_DAILY_LIST.GridAdvExColElement[mGRID_START_COL + mIDX_Column].HeaderElement[0].TL1_KR = iString.ISNull(mCOLUMN_DESC);

                //헤더 세팅.
                int vIDX_COL = 8;
                int vIDX_END_COL= vIDX_COL + 31;
                int vEXCEL_START_IDX = 3; 
                for (int c = vIDX_COL; c <= vIDX_END_COL; c++)
                {
                    if (iString.ISNull(pData.GridAdvExColElement[c].Visible) == "0")
                    {
                        mPrinting.XLSetCell(3, vEXCEL_START_IDX, "");
                        mPrinting.XLSetCell(4, vEXCEL_START_IDX, "");
                    }
                    else
                    {
                        vObject = pData.GridAdvExColElement[c].HeaderElement[0].Default;
                        if (vObject != null)
                        {
                            vString = string.Format("{0}", vObject);
                        }
                        else
                        {
                            vString = string.Empty;
                        }
                        mPrinting.XLSetCell(4, vEXCEL_START_IDX, vString);
                    }
                    vEXCEL_START_IDX++;
                }  
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        } 

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[8];
            pXLColumn = new int[8];

            string vDBColumn01 = "ACCOUNT_CODE";
            string vDBColumn02 = "ACCOUNT_DESC";
            string vDBColumn03 = "DR_AMOUNT";
            string vDBColumn04 = "CR_AMOUNT";
            string vDBColumn05 = "M_REFERENCE";
            string vDBColumn06 = "REMARK";
            string vDBColumn07 = "CUSTOMER_DESC";
            string vDBColumn08 = "DEPT_DESC";

            pDBColumn[0] = vDBColumn01;  //ACCOUNT_CODE
            pDBColumn[1] = vDBColumn02;  //ACCOUNT_DESC
            pDBColumn[2] = vDBColumn03;  //DR_AMOUNT
            pDBColumn[3] = vDBColumn04;  //CR_AMOUNT
            pDBColumn[4] = vDBColumn05;  //M_REFERENCE
            pDBColumn[5] = vDBColumn06;  //REMARK
            pDBColumn[6] = vDBColumn07;  //CUSTOMER_DESC
            pDBColumn[7] = vDBColumn08;  //DEPT_DESC

            int vXLColumn01 = 3;         //ACCOUNT_CODE
            int vXLColumn02 = 3;         //ACCOUNT_DESC
            int vXLColumn03 = 12;        //DR_AMOUNT
            int vXLColumn04 = 18;        //CR_AMOUNT
            int vXLColumn05 = 24;        //M_REFERENCE
            int vXLColumn06 = 24;        //REMARK
            int vXLColumn07 = 24;        //CUSTOMER_DESC
            int vXLColumn08 = 40;        //DEPT_DESC

            pXLColumn[0] = vXLColumn01;  //ACCOUNT_CODE
            pXLColumn[1] = vXLColumn02;  //ACCOUNT_DESC
            pXLColumn[2] = vXLColumn03;  //DR_AMOUNT
            pXLColumn[3] = vXLColumn04;  //CR_AMOUNT
            pXLColumn[4] = vXLColumn05;  //M_REFERENCE
            pXLColumn[5] = vXLColumn06;  //REMARK
            pXLColumn[6] = vXLColumn07;  //CUSTOMER_DESC
            pXLColumn[7] = vXLColumn08;  //DEPT_DESC
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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(int pRow, InfoSummit.Win.ControlAdv.ISGridAdvEx pData, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            object vObject;
            string vString= string.Empty;
            int vIDX_COL = 0;

            mCountLinePrinting++;

            mPrinting.XLActiveSheet(mTargetSheet); //셀에 문자를 넣기 위해 쉬트 선택

            try
            {
                //[성명]
                vIDX_COL = pData.GetColumnToIndex("P_NAME");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject); 
                }
                else
                {
                    vString = string.Empty; 
                }
                mPrinting.XLSetCell(vXLine, 1, vString);

                //[ITEM_TYPE_DESC]
                vIDX_COL = pData.GetColumnToIndex("ITEM_TYPE_DESC");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 2, vString);

                //[01]
                vIDX_COL = pData.GetColumnToIndex("D01");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 3, vString);

                //[02]
                vIDX_COL = pData.GetColumnToIndex("D02");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 4, vString);

                //[03]
                vIDX_COL = pData.GetColumnToIndex("D03");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 5, vString);

                //[04]
                vIDX_COL = pData.GetColumnToIndex("D04");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 6, vString);

                //[05]
                vIDX_COL = pData.GetColumnToIndex("D05");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 7, vString);

                //[06]
                vIDX_COL = pData.GetColumnToIndex("D06");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 8, vString);

                //[07]
                vIDX_COL = pData.GetColumnToIndex("D07");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 9, vString);

                //[08]
                vIDX_COL = pData.GetColumnToIndex("D08");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vString);

                //[09]
                vIDX_COL = pData.GetColumnToIndex("D09");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vString);

                //[10]
                vIDX_COL = pData.GetColumnToIndex("D10");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vString);

                //[11]
                vIDX_COL = pData.GetColumnToIndex("D11");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 13, vString);

                //[12]
                vIDX_COL = pData.GetColumnToIndex("D12");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vString);

                //[13]
                vIDX_COL = pData.GetColumnToIndex("D13");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vString);

                //[14]
                vIDX_COL = pData.GetColumnToIndex("D14");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 16, vString);

                //[15]
                vIDX_COL = pData.GetColumnToIndex("D15");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vString);

                //[16]
                vIDX_COL = pData.GetColumnToIndex("D16");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vString);

                //[17]
                vIDX_COL = pData.GetColumnToIndex("D17");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 19, vString);

                //[18]
                vIDX_COL = pData.GetColumnToIndex("D18");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vString);

                //[19]
                vIDX_COL = pData.GetColumnToIndex("D19");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 21, vString);

                //[20]
                vIDX_COL = pData.GetColumnToIndex("D20");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vString);

                //[21]
                vIDX_COL = pData.GetColumnToIndex("D21");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vString);

                //[22]
                vIDX_COL = pData.GetColumnToIndex("D22");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 24, vString);

                //[23]
                vIDX_COL = pData.GetColumnToIndex("D23");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vString);

                //[24]
                vIDX_COL = pData.GetColumnToIndex("D24");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 26, vString);

                //[25]
                vIDX_COL = pData.GetColumnToIndex("D25");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 27, vString);

                //[26]
                vIDX_COL = pData.GetColumnToIndex("D26");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vString);

                //[27]
                vIDX_COL = pData.GetColumnToIndex("D27");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vString);

                //[28]
                vIDX_COL = pData.GetColumnToIndex("D28");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 30, vString);

                //[29]
                vIDX_COL = pData.GetColumnToIndex("D29");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 31, vString);

                //[30]
                vIDX_COL = pData.GetColumnToIndex("D30");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 32, vString);

                //[31]
                vIDX_COL = pData.GetColumnToIndex("D31");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vString);

                //[total]
                vIDX_COL = pData.GetColumnToIndex("TOTAL");
                vObject = pData.GetCellValue(pRow, vIDX_COL);
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vString = string.Format("{0}", vObject);
                }
                else
                {
                    vString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 34, vString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            } 
            pPrintingLine = vXLine;

            return pPrintingLine;
         }
         
        #endregion;

        #region ----- Sum Write Methods -----

        private void SumWrite(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄// 
            int vPageRow = 54;
            int vLINE = 1;
            for (int r = 1; r <= mPageNumber; r++)
            {
                mPrinting.XLSetCell(vLINE, 33, string.Format("Page {0} of {1}", r, mPageNumber));
                vLINE = vLINE + vPageRow;

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
        }
         
        #endregion;

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pData)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            // 쉬트명 정의.
            mTargetSheet = "Sheet1";
            mSourceSheet1 = "Source1";
            mSourceSheet2 = "Source1"; 

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 35;
            mCopy_EndRow = 54;


            mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 52;  //최종 인쇄 라인.
            mCurrentRow = 5;
            int vPrintingLine = mCurrentRow;
            int vPerson_Row = 0;
            string vPERSON_NUM = string.Empty;
            
            try
            {
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                int vTotalRow = pData.RowCount;
                if (vTotalRow > 0)
                { 
                    int vCountRow = 0;
                    for(int r = 0; r < vTotalRow; r++)
                    {
                        vCountRow++;

                        vMessage = string.Format("{0}/{1}", vCountRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XlLine(r, pData, mCurrentRow);
                        vPrintingLine = vPrintingLine + 1;

                        if (vTotalRow == vCountRow)
                        {
                            //IsNewPage(vPrintingLine);
                            SumWrite(mCurrentRow);

                            //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);
                            //XlAllContentClear(mPrinting);
                        }
                        else
                        {
                            IsNewPage(vPrintingLine);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow + 2;
                                vPrintingLine = mDefaultPageRow + 1;

                            }
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

            return mPageNumber;
        }
         
        #endregion;

        #region ----- Last Page Number Compute Methods ----

        private void ComputeLastPageNumber(int pTotalRow)
        {
            int vRow = 0;
            mPageTotalNumber = 1;

            if (pTotalRow > 12)
            {
                vRow = pTotalRow - 12;
                mPageTotalNumber = vRow / 18;
                mPageTotalNumber = (vRow % 18) == 0 ? (mPageTotalNumber + 1) : (mPageTotalNumber + 2);
            }
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2,  mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }
            
        }

        private void IsNewPage_BSK(int pPrintingLine)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM);

                //XlAllContentClear(mPrinting);
            }
            else
            {
                mIsNewPage = false;
            }

        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
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
         
        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        public void PreView(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        public void Save(string pSaveFileName)
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
                mPrinting.XLDeleteSheet(mSourceSheet1);
                mPrinting.XLDeleteSheet(mSourceSheet2);
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