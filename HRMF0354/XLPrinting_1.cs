using System;

namespace HRMF0354
{
    public class XLPrinting_1
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART = 7;  //Line

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
        private int mIncrementCopyMAX = 52;  //복사되어질 행의 범위

        private int mCopyColumnSTART = 1; //복사되어  진 행 누적 수
        private int mCopyColumnEND = 77;  //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

        private int mTotal1ROW = 0;
        private int mIndex_DEPT_NAME = 0; //본부명
        private int mIndex_DEPT_CODE = 0; //본부코드

        private string mDepartmentCodeOLD = string.Empty;

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

        public XLPrinting_1(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        private void SetArray1(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[14];
            pXLColumn = new int[14];

            pGDColumn[0] = pTable.Columns.IndexOf("PERSON_NUM");        //사원번호
            pGDColumn[1] = pTable.Columns.IndexOf("NAME");              //성명
            pGDColumn[2] = pTable.Columns.IndexOf("POST_NAME");         //직위장
            pGDColumn[3] = pTable.Columns.IndexOf("FLOOR_NAME");        //작업장
            pGDColumn[4] = pTable.Columns.IndexOf("REMARK");            //업무내용
            pGDColumn[5] = pTable.Columns.IndexOf("PL_OT_START");       //출근[근무계획]
            pGDColumn[6] = pTable.Columns.IndexOf("PL_OT_END");         //퇴근[근무계획]
            pGDColumn[7] = pTable.Columns.IndexOf("OPEN_TIME");         //출근[근무시간]
            pGDColumn[8] = pTable.Columns.IndexOf("CLOSE_TIME");        //퇴근[근무시간]
            pGDColumn[9] = pTable.Columns.IndexOf("HOLIDAY_TIME");      //근무시간
            pGDColumn[10] = pTable.Columns.IndexOf("REAL_TIME");        //적용시간
            pGDColumn[11] = pTable.Columns.IndexOf("APPROVE_STATUS");   //승인
            pGDColumn[12] = pTable.Columns.IndexOf("APPROVED_PERSON");  //승인자
            pGDColumn[13] = pTable.Columns.IndexOf("DESCRIPTION");      //비고

            pXLColumn[0] = 1;    //사원번호
            pXLColumn[1] = 6;    //성명
            pXLColumn[2] = 11;   //직위
            pXLColumn[3] = 16;   //작업장
            pXLColumn[4] = 23;   //업무내용
            pXLColumn[5] = 37;   //출근[근무계획]
            pXLColumn[6] = 42;   //퇴근[근무계획]
            pXLColumn[7] = 47;   //출근[근무시간]
            pXLColumn[8] = 52;   //퇴근[근무시간]
            pXLColumn[9] = 57;   //근무시간(휴일)
            pXLColumn[10] = 61;  //적용시간(휴일)
            pXLColumn[11] = 65;  //승인
            pXLColumn[12] = 68;  //승인자
            pXLColumn[13] = 73;  //비고
        }

        #endregion;

        #region ----- Convert String Method ----

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
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
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
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {
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

        private void XLHeader(string pWorkDate, string pPrintingDateTime)
        {
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                vXLine = 3;
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, pWorkDate);

                vXLine = 51;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPrintingDateTime);
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

        private int XLLine(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");
                //사원번호
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
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
                //성명
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
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
                //직위
                vDBColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pRow[vDBColumnIndex];
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
                //작업장
                vDBColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pRow[vDBColumnIndex];
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
                //업무내용
                vDBColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pRow[vDBColumnIndex];
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
                //출근[근무계획]
                vDBColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pRow[vDBColumnIndex];
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
                //퇴근[근무계획]
                vDBColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pRow[vDBColumnIndex];
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
                //출근[근무시간]
                vDBColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pRow[vDBColumnIndex];
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
                //퇴근[근무시간]
                vDBColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pRow[vDBColumnIndex];
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
                //근무시간(휴일)
                vDBColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pRow[vDBColumnIndex];
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
                //적용시간(휴일)
                vDBColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pRow[vDBColumnIndex];
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
                //승인
                vDBColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pRow[vDBColumnIndex];
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
                //승인자
                vDBColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pRow[vDBColumnIndex];
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
                //비고
                vDBColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pRow[vDBColumnIndex];
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
                vXLine = vXLine + 2;
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

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, string pWorkDate)
        {
            mPageNumber = 0;
            string vMessage = string.Empty;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintingDateTime = string.Format("[인쇄 일시 : {0} {1}]", vPrintingDate, vPrintingTime);

            object vDEPT_NAME = null;
            object vDEPT_CODE = null;
            string vDepartmentCodeNEW = string.Empty;
            int vNewRow = 0;

            System.Data.DataRow vDataRow = null;

            int[] vGDColumn;
            int[] vXLColumn;

            int vPrintingLine = 0;

            try
            {
                mTotal1ROW = pAdapter.OraSelectData.Rows.Count;

                #region ----- Header Write ----

                XLHeader(pWorkDate, vPrintingDateTime);

                #endregion;

                #region ----- Line Write ----

                if (mTotal1ROW > 0)
                {
                    int vCountROW1 = 0;

                    mIndex_DEPT_NAME = pAdapter.OraSelectData.Columns.IndexOf("DEPT_NAME");
                    mIndex_DEPT_CODE = pAdapter.OraSelectData.Columns.IndexOf("DEPT_CODE");
                    vDEPT_NAME = pAdapter.OraSelectData.Rows[0][mIndex_DEPT_NAME];
                    vDEPT_CODE = pAdapter.OraSelectData.Rows[0][mIndex_DEPT_CODE];
                    vDepartmentCodeNEW = ConvertString(vDEPT_CODE);
                    mDepartmentCodeOLD = vDepartmentCodeNEW;

                    mCopyLineSUM = CopyAndPaste(mCopyLineSUM, vDEPT_NAME);

                    vPrintingLine = mPrintingLineSTART;

                    SetArray1(pAdapter.OraSelectData, out vGDColumn, out vXLColumn);

                    for (int vRow1 = 0; vRow1 < mTotal1ROW; vRow1++)
                    {
                        vCountROW1++;

                        vMessage = string.Format("Grid1 : {0}/{1}", vCountROW1, mTotal1ROW);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vDataRow = pAdapter.OraSelectData.Rows[vRow1];
                        vPrintingLine = XLLine(vDataRow, vPrintingLine, vGDColumn, vXLColumn);

                        if (mTotal1ROW == vCountROW1)
                        {
                            //마지막 행이면...
                        }
                        else
                        {
                            vNewRow = vRow1 + 1;
                            if (mTotal1ROW != vNewRow)
                            {
                                vDataRow = pAdapter.OraSelectData.Rows[vNewRow];
                            }
                            IsNewPage(vPrintingLine, vDataRow);
                            if (mIsNewPage == true)
                            {
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);
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

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine, System.Data.DataRow pRow)
        {
            object vObject = pRow[mIndex_DEPT_CODE];
            string vDepartmentCodeNEW = ConvertString(vObject);

            int vPrintingLineEND = mCopyLineSUM - 4; //1~52: mCopyLineSUM=53에서 내용이 출력되는 행이 49 이므로, 4을 빼면 된다
            if (mDepartmentCodeOLD != vDepartmentCodeNEW)
            {
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pRow[mIndex_DEPT_NAME]);
                mIsNewPage = true;
                mDepartmentCodeOLD = vDepartmentCodeNEW;
            }
            else if (vPrintingLineEND < pPrintingLine)
            {
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pRow[mIndex_DEPT_NAME]);
                mIsNewPage = true;
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //첫번째 페이지 복사
        private int CopyAndPaste(int pCopySumPrintingLine, object pDEPT_NAME)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            mPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = mPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            //본부명
            string vDepartmentName = string.Format("본부명 : {0}", pDEPT_NAME);
            int vDrawRow = vCopyPrintingRowSTART + 2;
            mPrinting.XLSetCell((vDrawRow + 0), 2, vDepartmentName);

            mPageNumber++; //페이지 번호

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

        public void SAVE(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}