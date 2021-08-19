using System;
using ISCommonUtil;

namespace HRMF0364
{
    public class XLPrinting
    {
        
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SourceTab1";
        //private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;   //���� �μ� ����.
        //private int m1stPrintingLastRow = 0;
        private int mCurrentRow = 0;        //���� �μ�Ǵ� row ��ġ.
        //private int mDefaultEndPageRow = 1; // ������ ������ PageCount �⺻��.
        private int mDefaultPageRow = 4;    // ������ ������ PageCount �⺻��.

        //private string[] mGridColumn; 

        //Copy�Ҷ� �����ؾ��� ���� �� ��ġ ���
        private int[] mRowMerge = new int[8] { -1, -1, -1, -1, -1, -1, -1, -1 };
        private int mCountRow = 0; //�����ؾ��� ���� �� ��ġ Count 

        private object mPringingDateTime = string.Empty;
        private object mReq_Person_Name = string.Empty;

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

        //public int PrintingLineMAX
        //{
        //    set
        //    {
        //        mPrintingLineMAX = value;
        //    }
        //}

        //public int IncrementCopyMAX
        //{
        //    set
        //    {
        //        mIncrementCopyMAX = value;
        //    }
        //}

        //public int PositionPrintLineSTART
        //{
        //    set
        //    {
        //        mPositionPrintLineSTART = value;
        //    }
        //}

        //public int CopySumPrintingLine
        //{
        //    set
        //    {
        //        mCopySumPrintingLine = value;
        //    }
        //}

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterfaceAdv, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterfaceAdv;
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

        private void XlAllLineClear(object pCORP_NAME)
        {
            object vObject = null;

            mPrinting.XLActiveSheet(mTargetSheet);

            int vStartRow = mCurrentRow;
            int vStartCol = mCopy_StartCol + 1;
            int vEndRow = mCopyLineSUM;
            int vEndCol = mCopy_EndCol - 1;

            mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, vObject);
            mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
            mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
            mPrinting.XL_LineDraw_Top(vStartRow, vStartCol, vEndCol, 2);  //���� ������ �־.

            mPrinting.XLCellMerge(mCurrentRow, 51, mCurrentRow, vEndCol, true);
            mPrinting.XLCellAlignmentHorizontal(mCurrentRow, 51, mCurrentRow, vEndCol, "R");
            mPrinting.XLSetCell(mCurrentRow, 51, pCORP_NAME);
        }

        private void RateLineClear(int pPrintingLine, int pCopyPrintingRowSTART, int pCopyPrintingRowEnd)
        {

            int vStartRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;
            int vStartCol = mCopy_StartCol + 1;
            int vEndRow = pCopyPrintingRowEnd - 1;
            int vEndCol = mCopy_EndCol;
            int vDrawRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;

            mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
            mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
            mPrinting.XL_LineDraw_Top(vDrawRow, vStartCol, vEndCol, 2);
        }

        #endregion;

        #region ----- Cell Merge Methods ----

        private void CellMerge(int pCopySumPrintingLine, int pCountRow, int[] pRowMerge)
        {            
            //int vXLine = 0;
            int vCountRowMerge = pRowMerge.Length;

            try
            {
                for (int vCount = 0; vCount < vCountRowMerge; vCount++)
                {
                    if (pRowMerge[vCount] == 1)
                    {
                        //vXLine = pCopySumPrintingLine + mPositionPrintLineSTART + (vCount * 4);
                        //int vStartRow = vXLine - 1;
                        //int vStartCol = mXLColumn[1];
                        //int vEndRow = vXLine + 2;
                        //int vEndCol = mXLColumn[3] - 1;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                        //vXLine = pCopySumPrintingLine + mPositionPrintLineSTART + (vCount * 4);
                        //int vStartRow = vXLine - 1;
                        //int vStartCol = mXLColumn[1];
                        //int vEndRow = vXLine;
                        //int vEndCol = mXLColumn[3] - 1;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);

                        //vStartRow = vXLine + 1;
                        //vEndRow = vXLine + 2;

                        //mPrinting.XLCellMerge(vStartRow, vStartCol, vEndRow, vEndCol, false);
                    }

                    mRowMerge[vCount] = -1;
                }

                mCountRow = 0; //�����ؾ��� ���� �� ��ġ Count, 0���� Set
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }        
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(object pUserName, object pPrintingDateTime, object pDepartment_NAME)
        { 
            try
            {               
                mPrinting.XLActiveSheet(mSourceSheet1); //���� ���ڸ� �ֱ� ���� ��Ʈ ����
                
                //����� 
                if (iConv.ISNull(pUserName) != string.Empty)
                {
                    mPrinting.XLSetCell(34, 1, pUserName);
                } 
                 
                //�۾��� 
                mPrinting.XLSetCell(2, 24, pDepartment_NAME);
                 
                //������� 
                if (iConv.ISNull(pPrintingDateTime) != string.Empty)
                {
                    mPrinting.XLSetCell(34, 53, string.Format("{0:yyyy-MM-dd hh:mm:dd}", pPrintingDateTime));
                }
                else
                {
                    mPrinting.XLSetCell(34, 53, null);
                } 
            }
            catch (System.Exception ex)
            {
                mAppInterface.OnAppMessage(ex.Message);

                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Array Set ----

        private void SetArray(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGridColumn)
        {
            pGridColumn = new int[81];

            pGridColumn[0] = pGrid.GetColumnToIndex("DEPT_NAME");                   //�μ�
            pGridColumn[1] = pGrid.GetColumnToIndex("POST_NAME");                   //����
            pGridColumn[2] = pGrid.GetColumnToIndex("PERSON_NUM");                  //�����ȣ
            pGridColumn[3] = pGrid.GetColumnToIndex("NAME");                        //����
            pGridColumn[4] = pGrid.GetColumnToIndex("PAY_TYPE_NAME");               //�޿�����
            pGridColumn[5] = pGrid.GetColumnToIndex("JOIN_DATE");                   //�Ի�����.
            pGridColumn[6] = pGrid.GetColumnToIndex("RETIRE_DATE");                 //�������
            pGridColumn[7] = pGrid.GetColumnToIndex("LONG_YEAR");                   //�ټӳ��
            pGridColumn[8] = pGrid.GetColumnToIndex("LONG_MONTH");                  //�ټӿ���
            pGridColumn[9] = pGrid.GetColumnToIndex("BASIC_BASE_AMOUNT");           //�⺻��
            pGridColumn[10] = pGrid.GetColumnToIndex("GENERAL_HOURLY_PAY_AMOUNT");  //���ñ�

            pGridColumn[11] = pGrid.GetColumnToIndex("PAY_DAY");                    //�޿��ϼ�
            pGridColumn[12] = pGrid.GetColumnToIndex("A01");                        //�⺻��
            pGridColumn[13] = pGrid.GetColumnToIndex("LATE_TIME");                  //����/����
            pGridColumn[14] = pGrid.GetColumnToIndex("A17");                        //���°����ݾ�
            pGridColumn[15] = pGrid.GetColumnToIndex("OVER_TIME");                  //����ð�
            pGridColumn[16] = pGrid.GetColumnToIndex("A12");                        //����ݾ�

            pGridColumn[17] = pGrid.GetColumnToIndex("HOLY_1_TIME");                //���ϱٷνð�
            pGridColumn[18] = pGrid.GetColumnToIndex("A14");                        //���ϱٷαݾ�
            pGridColumn[19] = pGrid.GetColumnToIndex("HOLY_0_TIME");                //���ٷνð�
            pGridColumn[20] = pGrid.GetColumnToIndex("A20");                        //���ٷαݾ�
            pGridColumn[21] = pGrid.GetColumnToIndex("NIGHT_BONUS");                //�߰������ð�
            pGridColumn[22] = pGrid.GetColumnToIndex("A13");                        //�߰������ݾ�

            pGridColumn[23] = pGrid.GetColumnToIndex("A02");                        //��å����
            pGridColumn[24] = pGrid.GetColumnToIndex("A11");                        //�ð��ܼ���
            pGridColumn[25] = pGrid.GetColumnToIndex("A25");                        //����������
            pGridColumn[26] = pGrid.GetColumnToIndex("A30");                        //�����
            pGridColumn[27] = pGrid.GetColumnToIndex("A32");                        //���ٹ�����
            pGridColumn[28] = pGrid.GetColumnToIndex("A22");                        //��ٰ���
            pGridColumn[29] = pGrid.GetColumnToIndex("A24");                        //��������
            pGridColumn[30] = pGrid.GetColumnToIndex("A09");                        //�󿩱�
            pGridColumn[31] = pGrid.GetColumnToIndex("A07");                        //��Ÿ����
            pGridColumn[32] = pGrid.GetColumnToIndex("A28");                        //���ټ���
            pGridColumn[33] = pGrid.GetColumnToIndex("A27");                        //ö�߼���
            pGridColumn[34] = pGrid.GetColumnToIndex("A37");                        //
            pGridColumn[35] = pGrid.GetColumnToIndex("A38");                        //
            pGridColumn[36] = pGrid.GetColumnToIndex("A37");                        //
            pGridColumn[37] = pGrid.GetColumnToIndex("A39");                        //
            pGridColumn[38] = pGrid.GetColumnToIndex("A38");                        //
            pGridColumn[39] = pGrid.GetColumnToIndex("ETC_SUM");                    //��Ÿ�����հ�
            pGridColumn[40] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          //�������հ�

            pGridColumn[41] = pGrid.GetColumnToIndex("D01");                        //�ҵ漼
            pGridColumn[42] = pGrid.GetColumnToIndex("D02");                        //�ֹμ�            
            pGridColumn[43] = pGrid.GetColumnToIndex("D03");                        //���ο���
            pGridColumn[44] = pGrid.GetColumnToIndex("D04");                        //��뺸��
            pGridColumn[45] = pGrid.GetColumnToIndex("D05");                        //�ǰ�����
            pGridColumn[46] = pGrid.GetColumnToIndex("D06");                        //����纸��
            pGridColumn[47] = pGrid.GetColumnToIndex("D07");                        //�ǰ����������
            pGridColumn[48] = pGrid.GetColumnToIndex("D08");                        //��纸�������
            pGridColumn[49] = pGrid.GetColumnToIndex("D09");                        //���ұ�
            pGridColumn[50] = pGrid.GetColumnToIndex("D10");                        //���������
            pGridColumn[51] = pGrid.GetColumnToIndex("D11");                        //�Ǻ���
            pGridColumn[52] = pGrid.GetColumnToIndex("D12");                        //������߱޺�
            pGridColumn[53] = pGrid.GetColumnToIndex("D13");                        //���νſ뺸��
            pGridColumn[54] = pGrid.GetColumnToIndex("D14");                        //��Ÿ����
            pGridColumn[55] = pGrid.GetColumnToIndex("D15");                        //����ҵ漼
            pGridColumn[56] = pGrid.GetColumnToIndex("D16");                        //�����ֹμ�
            pGridColumn[57] = pGrid.GetColumnToIndex("D17");                        //�����Ư��
            pGridColumn[58] = pGrid.GetColumnToIndex("D18");                        // 
            pGridColumn[59] = pGrid.GetColumnToIndex("D19");                        //���з����� 
            pGridColumn[60] = pGrid.GetColumnToIndex("D20");                        // 

            pGridColumn[61] = pGrid.GetColumnToIndex("D21");                        // 
            pGridColumn[62] = pGrid.GetColumnToIndex("D22");                        // 
            pGridColumn[63] = pGrid.GetColumnToIndex("D23");                        // 
            pGridColumn[64] = pGrid.GetColumnToIndex("D24");                        // 
            pGridColumn[65] = pGrid.GetColumnToIndex("D25");                        //��������ҵ漼
            pGridColumn[66] = pGrid.GetColumnToIndex("D26");                        //���������ֹμ�
            pGridColumn[67] = pGrid.GetColumnToIndex("D27");                        //���������Ư��
            pGridColumn[68] = pGrid.GetColumnToIndex("D28");                        //����ȸ��
            pGridColumn[69] = pGrid.GetColumnToIndex("D29");                        //  

            pGridColumn[70] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");             //�Ѱ����� 
            pGridColumn[71] = pGrid.GetColumnToIndex("REAL_AMOUNT");                //�����޾�

            pGridColumn[72] = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");              //�������
            pGridColumn[73] = pGrid.GetColumnToIndex("DUTY_30");                    //����
            pGridColumn[74] = pGrid.GetColumnToIndex("TOT_DED_COUNT");              //�̱ٹ�
            pGridColumn[75] = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");             //����
            pGridColumn[76] = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");           //������
            pGridColumn[77] = pGrid.GetColumnToIndex("HOLY_1_COUNT");               //����
            pGridColumn[78] = pGrid.GetColumnToIndex("HOLY_0_COUNT");               //����
            pGridColumn[79] = pGrid.GetColumnToIndex("DEPT_CODE");                  //�μ��ڵ�
            pGridColumn[80] = pGrid.GetColumnToIndex("SUMMARY_FLAG");               //�հ迩��
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
                mAppInterface.OnAppMessage(mMessageError);
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
                mAppInterface.OnAppMessage(mMessageError);
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
                mAppInterface.OnAppMessage(mMessageError);
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

        #region ----- Line Print ----- 
         
        private int XlLine(System.Data.DataRow pRow, int pPrintingLine)
        {
            int vXLine = pPrintingLine; //������ ������ ǥ�õǴ� �� ��ȣ

            object vGetValue = null;  

            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            //string vSUMMARY_FLAG = "N";

            bool IsConvert = false;  
            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //[����] 
                vGetValue = pRow["NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {
                    
                }
                else
                {
                    vConvertString = string.Empty;
                } 
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //[����] 
                vGetValue = pRow["POST_NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty; 
                }
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //[������] 
                vGetValue = pRow["JOB_CATEGORY_NAME"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 10, vConvertString);

                //[�ٹ�����] 
                vGetValue = pRow["WORK_DATE"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 13, vConvertString);

                //[����] 
                vGetValue = pRow["DANGJIK_YN"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 18, vConvertString);

                //[ö��] 
                vGetValue = pRow["ALL_NIGHT_YN"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vConvertString);

                //[�ٹ���-����] 
                vGetValue = pRow["BEFORE_TIME_START"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vConvertString);

                //[�ٹ���-����] 
                vGetValue = pRow["BEFORE_TIME_END"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 25, vConvertString);

                //[�ٹ���-����] 
                vGetValue = pRow["AFTER_OT_START"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 28, vConvertString);

                //[�ٹ���-����] 
                vGetValue = pRow["AFTER_OT_END"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vConvertString);

                //[����] 
                vGetValue = pRow["BREAKFAST_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 42, vConvertString);

                //[�߽�] 
                vGetValue = pRow["LUNCH_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 44, vConvertString);

                //[����] 
                vGetValue = pRow["DINNER_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 46, vConvertString);

                //[�߽�] 
                vGetValue = pRow["MIDNIGHT_FLAG"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 48, vConvertString);

                //[����]
                vGetValue = pRow["DESCRIPTION"];
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 50, vConvertString); 

                vXLine = vXLine + 1;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
            } 
            return vXLine;
        }

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int XLWirteMain(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter
                                , object pLocal_DATE
                                , object pReq_Person_Name)
        {
            string vMessage = string.Empty;
            mIsNewPage = false; 

            //�ʱ�ȭ//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 62;
            mCopy_EndRow = 34;

            //mDefaultEndPageRow = 1;
            mDefaultPageRow = 8;    // ������ ������ PageCount �⺻��.
            mPrintingLastRow = 33;  //���� �μ� ����.
            //m1stPrintingLastRow = 40;

            mCurrentRow = 8;
            mCopyLineSUM = 1;

            int vTotalRow = 0;
            int vPageRowCount = 0;  //�μ��� �ش� ���� ���� ����. 
            int vCurrRow = 0;

            mPringingDateTime = pLocal_DATE;

            string vDEPT_CODE = string.Empty; 
            try
            {
                vTotalRow = pAdapter.CurrentRows.Count;
                //TotalPage(pGrid);

                if (vTotalRow > 0)
                {
                    //�迭 ����. 
                    vPageRowCount = mCurrentRow - 1;  
                    foreach(System.Data.DataRow vRow in pAdapter.CurrentRows)
                    {
                        vCurrRow++;
                        vMessage = string.Format("Row : {0} / {1}", vCurrRow, vTotalRow);
                        mAppInterface.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        if (vCurrRow == 1)
                        {
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pReq_Person_Name, vRow["FLOOR_NAME"]);
                        }
                        //if (vRow == 0)
                        //{
                        //    //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pGrid, vRow, vDEPT_NAME);
                        //    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME); 
                        //}
                        //else if (vDEPT_CODE != iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[79])) && mIsNewPage == false)
                        //{
                        //    //XlAllLineClear(pCorporationName);
                        //    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME);
                        //    //�����μ� �� �̹Ƿ� ������ROW�� +4�� ����.
                        //    mCurrentRow = mCurrentRow + (mCopy_EndRow - (vPageRowCount + 4)) + mDefaultPageRow;  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                        //    vPageRowCount = mDefaultPageRow - 4;
                        //}

                        mCurrentRow = XlLine(vRow, mCurrentRow);
                        vPageRowCount = vPageRowCount + 1;
                        
                        IsNewPage(mPrinting, vPageRowCount, vDEPT_CODE, iConv.ISNull(vRow["FLOOR_CODE"]), pReq_Person_Name, vRow["FLOOR_NAME"]);   // ���ο� ������ üũ �� ����.
                        if (mIsNewPage == true)
                        {
                            //�μ� �� �̹Ƿ� ���� ������ROW�� -4�� ����.
                            mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 1) + mDefaultPageRow;  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                            vPageRowCount = mDefaultPageRow - 1;
                        }
                        vDEPT_CODE = iConv.ISNull(vRow["FLOOR_CODE"]);

                        //if (vRow == vTotalRow -1)
                        //{
                        //    // ������ ������ �̸� ó���� ���� ���
                        //    // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                        //    SumWrite(mCurrentRow);      //�հ�.
                        //    if (vPageRowCount != mPrintingLastRow)
                        //    {
                        //        //������ROW�� ������ �μ��ϰ� �ٸ��� ���� ���� CLEAR
                        //        XlAllLineClear(pCorporationName);
                        //    }
                        //}
                        //else
                        //{
                        //    IsNewPage(vPageRowCount, false, vDEPT_NAME);   // ���ο� ������ üũ �� ����.
                        //    if (mIsNewPage == true)
                        //    {
                        //        //�μ� �� �̹Ƿ� ���� ������ROW�� -4�� ����.
                        //        mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 4) + mDefaultPageRow;  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                        //        vPageRowCount = mDefaultPageRow - 4;
                        //    }
                        //} 
                    }

                    //mPrinting.XLDeleteSheet(mSourceSheet1); 
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

        #region ----- TOTAL AMOUNT Write Method -----

        private void SumWrite(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber �μ�//
            int vPageCount = 45;
            int vLINE = 0;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageCount * (r - 1);
                mPrinting.XLSetCell((vLINE + 4), 56, string.Format("Page {0} of {1}", r, mPageNumber));

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

            ////�հ� �μ�//
            //vLINE = mPageNumber * mCopy_EndRow;
            //vLINE = vLINE - 1;
            ////mPrinting.XLSetCell(vLINE, 1, "SUM");
            //string vAmount = string.Empty;

            ////[�հ�]
            //if (mPageNumber == 1)
            //{
            //    vLINE = 31;
            //    mPrinting.XLSetCell(vLINE, 1, "[��    ��]");

            //    //BACK COLOR.
            //    mPrinting.XLCellColorBrush(vLINE, 8, vLINE, 15, System.Drawing.Color.Silver);

            //    //��ȹ�հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 8, vAmount);

            //    //�����հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 11, vAmount);

            //    //�����հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 14, vAmount);

            //    //XlLineClear(pPrintingLine);

            //}
            //else
            //{
            //    mPrinting.XLSetCell(vLINE, 1, "[��    ��]");

            //    //BACK COLOR.
            //    mPrinting.XLCellColorBrush(vLINE, 8, vLINE, 15, System.Drawing.Color.Silver);

            //    //��ȹ�հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 8, vAmount);

            //    //�����հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 11, vAmount);

            //    //�����հ�
            //    vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
            //    mPrinting.XLSetCell(vLINE, 14, vAmount);

            //    //XlLineClear(pPrintingLine);
            //}
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, object pDEPT_CODE, object pReq_Person_Name, object pDEPT_NAME)
        {
            if (mPrintingLastRow == pPrintingLine)
            {
                mIsNewPage = true;                
                //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true; 
                //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pPrintingLine, pGrid, pRow);  
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mCopyLineSUM, pPrintingLine, pGrid, pRow); 
            }
            else
            {
                mIsNewPage = false;
            }
        }

        private void IsNewPage(XL.XLPrint pPrinting, int pPrintingLine, object pOLD_DEPT_CODE, object pDEPT_CODE, object pReq_Person_Name, object pDEPT_NAME)
        {
            if (mPrintingLastRow < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(pPrinting, mCopyLineSUM, pReq_Person_Name, pDEPT_NAME);
            }
            else if (iConv.ISNull(pOLD_DEPT_CODE) != string.Empty && iConv.ISNull(pOLD_DEPT_CODE) != iConv.ISNull(pDEPT_CODE))
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(pPrinting, mCopyLineSUM, pReq_Person_Name, pDEPT_NAME); 
            } 
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, object pReq_Person_Name, object pDEPT_NAME)
        {
            mPageNumber++; //������ ��ȣ

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //�� �Լ��� ȣ�� ���� ������ �׸������� XL Sheet�� Insert ���� �ʴ´�.

            HeaderWrite(pReq_Person_Name, mPringingDateTime, pDEPT_NAME);
            //DepartmentName(pGrid, pRow);

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            int vCopyPrintingRowSTART = pCopySumPrintingLine;

            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, object pDEPT_NAME)
        {
            mPageNumber++; //������ ��ȣ

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //�� �Լ��� ȣ�� ���� ������ �׸������� XL Sheet�� Insert ���� �ʴ´�.

            //HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);            
            //DepartmentName(pGrid, pRow);

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            int vCopyPrintingRowSTART = pCopySumPrintingLine;

            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
        }

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(int pCopySumPrintingLine, int pPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        {
            int vPrintHeaderColumnSTART = mCopy_StartCol; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = mCopy_EndCol;     //����Ǿ��� ��Ʈ�� ��, ���῭

            mPageNumber++;
            //mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
            //HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mCorporationName);
            //DepartmentName(pGrid, pRow);

            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            mPrinting.XLActiveSheet(mSourceSheet1);
            object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mCopy_EndRow, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            //��ü
            int vDrawRow = (pPrintingLine + vCopyPrintingRowSTART) - 1;
            //mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

            CellMerge(pCopySumPrintingLine, mCountRow, mRowMerge);

            RateLineClear(pPrintingLine, vCopyPrintingRowSTART, vCopyPrintingRowEnd);

            return vCopySumPrintingLine;
        }

        ////[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        //private int CopyAndPaste_1(int pCopySumPrintingLine, int pPrintingLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow)
        //{
        //    int vPrintHeaderColumnSTART = mCopy_StartCol; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
        //    int vPrintHeaderColumnEND = mCopy_EndCol;     //����Ǿ��� ��Ʈ�� ��, ���῭

        //    mCountPage++;
        //    mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
        //    HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mPageString, mCorporationName);
        //    DepartmentName(pGrid, pRow);

        //    int vCopySumPrintingLine = pCopySumPrintingLine;

        //    int vCopyPrintingRowSTART = vCopySumPrintingLine;
        //    vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
        //    int vCopyPrintingRowEnd = vCopySumPrintingLine;
        //    mPrinting.XLActiveSheet("SourceTab1");
        //    object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mIncrementCopyMAX, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    mPrinting.XLActiveSheet("Destination");
        //    object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

        //    //��ü
        //    int vDrawRow = (pPrintingLine + vCopyPrintingRowSTART) - 1;
        //    mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

        //    CellMerge(pCopySumPrintingLine, mCountRow, mRowMerge);

        //    RateLineClear(pPrintingLine, vCopyPrintingRowSTART, vCopyPrintingRowEnd);

        //    return vCopySumPrintingLine;
        //}

        #endregion;

        //#region ----- Excel Rate Line Clear Method ----

        //private void RateLineClear(int pPrintingLine, int pCopyPrintingRowSTART, int pCopyPrintingRowEnd)
        //{

        //    int vStartRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;
        //    int vStartCol = mCopyColumnSTART + 1;
        //    int vEndRow = pCopyPrintingRowEnd - 1;
        //    int vEndCol = mCopyColumnEND;
        //    int vDrawRow = (pPrintingLine + pCopyPrintingRowSTART) - 1;

        //    mPrinting.XL_LineClearALL(vStartRow, vStartCol, vEndRow, vEndCol);
        //    mPrinting.XLCellColorBrush(vStartRow, vStartCol, vEndRow, vEndCol, System.Drawing.Color.White);
        //    mPrinting.XL_LineDraw_Top(vDrawRow, vStartCol, vEndCol, 2);
        //}

        //#endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPrinting(pPageSTART, pPageEND);
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }
            mPrinting.XLSave(pSaveFileName);

            //��ȣ�� �ּ�
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- Save Methods ----

        public void PDF_Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }
            mPrinting.XLSaveAs_PDF(pSaveFileName);

            //��ȣ�� �ּ�
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// ���������� ������Ʈ �����ϱ� ���� ������Ʈ�� ����ϰ� ��Ʈ�� �����Ѵ�.

            int vXLRow = 31; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLCol = 40;

            try
            { // ������ �����ؼ� Ÿ�� �� ������ ����.(
                mPrinting.XLActiveSheet(pActiveSheet);
                mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessage(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;
                
    }
}