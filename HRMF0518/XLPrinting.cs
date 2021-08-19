using System;
using ISCommonUtil;

namespace HRMF0518
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
        private string mTargetSheet = "Destination";
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

        private object mCorporationName = string.Empty;
        private object mUserName = string.Empty;
        private object mYYYYMM = string.Empty;
        private object mWageTypeName = string.Empty;
        private object mDepartmentName = string.Empty;
        private object mPringingDateTime = string.Empty;

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

        public void HeaderWrite(object pUserName, object pPrintingDateTime, object pYYYYMM, object pWageTypeName, object pDepartment_NAME, object pCorporationName)
        { 
            try
            {
                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(2, 2);    //Title
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(4, 6);    //�����
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(5, 6);    //�޿�����
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 19);   //�μ�
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(4, 56);   //������
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(5, 56);   //�������
                System.Drawing.Point vCellPoint07 = new System.Drawing.Point(44, 51);  //��ü

                mPrinting.XLActiveSheet(mSourceSheet1); //���� ���ڸ� �ֱ� ���� ��Ʈ ����

                //Title 
                if (iConv.ISNull(pYYYYMM) != string.Empty)
                {
                    string vYear = iConv.ISNull(pYYYYMM).Substring(0, 4);
                    string vMonth = iConv.ISNull(pYYYYMM).Substring(5, 2);
                    string vTitle = string.Format("{0}�� {1}�� {2} ����", vYear, vMonth, pWageTypeName);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vTitle);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, null);
                }

                //����� 
                if (iConv.ISNull(pUserName) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, pUserName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, null);
                }

                //�޿����� 
                if (iConv.ISNull(pWageTypeName) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, pWageTypeName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, "��ü");
                }

                //�μ� 
                mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, pDepartment_NAME);

                ////������ 
                //if (iConv.ISNull(pPageString) != string.Empty)
                //{
                //    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, pPageString);
                //}
                //else
                //{
                //    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, null);
                //}

                //������� 
                if (iConv.ISNull(pPrintingDateTime) != string.Empty)
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, string.Format("{0:yyyy-MM-dd hh:mm:dd}", pPrintingDateTime));
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, null);
                }

                //��ü
                mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, pCorporationName); 
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
            pGridColumn = new int[84];

            pGridColumn[0] = pGrid.GetColumnToIndex("DEPT_NAME");                   //�μ�
            pGridColumn[1] = pGrid.GetColumnToIndex("POST_NAME");                   //����
            pGridColumn[2] = pGrid.GetColumnToIndex("PERSON_NUM");                  //�����ȣ
            pGridColumn[3] = pGrid.GetColumnToIndex("NAME");                        //����
            pGridColumn[4] = pGrid.GetColumnToIndex("PAY_TYPE_NAME");               //�޿�����
            pGridColumn[5] = pGrid.GetColumnToIndex("JOIN_DATE");                   //�Ի�����.
            pGridColumn[6] = pGrid.GetColumnToIndex("RETIRE_DATE");                 //�������
            pGridColumn[7] = pGrid.GetColumnToIndex("LONG_YEAR");                   //�ټӳ��
            pGridColumn[8] = pGrid.GetColumnToIndex("LONG_MONTH");                  //�ټӿ���
            pGridColumn[9] = pGrid.GetColumnToIndex("BAISC_AMOUNT");             //������
            pGridColumn[10] = pGrid.GetColumnToIndex("GENERAL_HOURLY_PAY_AMOUNT");  //���ñ�

            pGridColumn[11] = pGrid.GetColumnToIndex("PAY_DAY");                    //�޿��ϼ�
            pGridColumn[12] = pGrid.GetColumnToIndex("A01");                        //�⺻��
            pGridColumn[13] = pGrid.GetColumnToIndex("LATE_TIME");                  //����/����
            pGridColumn[14] = pGrid.GetColumnToIndex("A07");                        //��Ÿ����
            pGridColumn[15] = pGrid.GetColumnToIndex("OVER_TIME");                  //����ð�
            pGridColumn[16] = pGrid.GetColumnToIndex("A29");                        //�ɾ߼���

            pGridColumn[17] = pGrid.GetColumnToIndex("HOLY_1_TIME");                //���ϱٷνð�
            pGridColumn[18] = pGrid.GetColumnToIndex("A14");                        //���ϱٷαݾ�
            pGridColumn[19] = pGrid.GetColumnToIndex("HOLY_0_TIME");                //���Ͽ���
            pGridColumn[20] = pGrid.GetColumnToIndex("A20");                        //���Ͽ���ݾ�
            pGridColumn[21] = pGrid.GetColumnToIndex("NIGHT_BONUS");                //�߰������ð�
            pGridColumn[22] = pGrid.GetColumnToIndex("PRODUCTION_SUM");           //���������

            pGridColumn[23] = pGrid.GetColumnToIndex("A02");                        //����������
            pGridColumn[24] = pGrid.GetColumnToIndex("A30");                        //�ޱټ���
            pGridColumn[25] = pGrid.GetColumnToIndex("A38");                        //���޼���
            pGridColumn[26] = pGrid.GetColumnToIndex("A06");                        //�ڰݼ���
            pGridColumn[27] = pGrid.GetColumnToIndex("AMOUNT1");              //����ٷμ��� + �ɾ߼���
            pGridColumn[28] = pGrid.GetColumnToIndex("A34");                        //��Ÿ����2
            pGridColumn[29] = pGrid.GetColumnToIndex("A24");                        //��������
            pGridColumn[30] = pGrid.GetColumnToIndex("A09");                        //�󿩱�
            pGridColumn[31] = pGrid.GetColumnToIndex("A08");                        //���м���
            pGridColumn[32] = pGrid.GetColumnToIndex("A28");                        //���ټ���
            pGridColumn[33] = pGrid.GetColumnToIndex("A27");                        //ö�߼���
            pGridColumn[34] = pGrid.GetColumnToIndex("A04");                        //���޼���
            pGridColumn[35] = pGrid.GetColumnToIndex("A17");                        //���°���
            pGridColumn[36] = pGrid.GetColumnToIndex("A03");                        //�޿��ұ޺�
            pGridColumn[37] = pGrid.GetColumnToIndex("A33");                        //���ܱٷμ���
            pGridColumn[38] = pGrid.GetColumnToIndex("A13");                        //�߰��ٹ� 
            pGridColumn[39] = pGrid.GetColumnToIndex("A08");                       //���м���
            pGridColumn[81] = pGrid.GetColumnToIndex("A35");                       //����������
            pGridColumn[82] = pGrid.GetColumnToIndex("A36");                       //���庹����
            pGridColumn[83] = pGrid.GetColumnToIndex("A05");                       //���庹����
            pGridColumn[40] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          //�������հ�

            pGridColumn[41] = pGrid.GetColumnToIndex("D01");                        //�ҵ漼
            pGridColumn[42] = pGrid.GetColumnToIndex("D02");                        //�ֹμ�            
            pGridColumn[43] = pGrid.GetColumnToIndex("D03");                        //���ο���
            pGridColumn[44] = pGrid.GetColumnToIndex("D04");                        //��뺸��
            pGridColumn[45] = pGrid.GetColumnToIndex("AMOUNT5");               //�ǰ�����
            pGridColumn[46] = pGrid.GetColumnToIndex("D21");                        //����纸��  //���ݻ�ȯ
            pGridColumn[47] = pGrid.GetColumnToIndex("D07");                        //�ǰ����������  
            pGridColumn[48] = pGrid.GetColumnToIndex("D08");                        //��纸�������
            pGridColumn[49] = pGrid.GetColumnToIndex("D22");                        //���ұ�  //���ڳ���
            pGridColumn[50] = pGrid.GetColumnToIndex("D10");                        //��������� 
            pGridColumn[51] = pGrid.GetColumnToIndex("D11");                        //�Ǻ���
            pGridColumn[52] = pGrid.GetColumnToIndex("D12");                        //������߱޺�
            pGridColumn[53] = pGrid.GetColumnToIndex("D13");                        //���νſ뺸��
            pGridColumn[54] = pGrid.GetColumnToIndex("ETC_DED_TOTAL");       //��Ÿ����
            pGridColumn[55] = pGrid.GetColumnToIndex("D32");                        //����ҵ漼  //���ڱ�
            pGridColumn[56] = pGrid.GetColumnToIndex("D16");                        //�����ֹμ� //
            pGridColumn[57] = pGrid.GetColumnToIndex("D17");                        //�����Ư��
            pGridColumn[58] = pGrid.GetColumnToIndex("D18");                        // 
            pGridColumn[59] = pGrid.GetColumnToIndex("D19");                        //���з����� 
            pGridColumn[60] = pGrid.GetColumnToIndex("D20");                        // 

            pGridColumn[61] = pGrid.GetColumnToIndex("D06");                        // 
            pGridColumn[62] = pGrid.GetColumnToIndex("D22");                        // ���������
            pGridColumn[63] = pGrid.GetColumnToIndex("D23");                        // 
            pGridColumn[64] = pGrid.GetColumnToIndex("D24");                        // 
            pGridColumn[65] = pGrid.GetColumnToIndex("AMOUNT2");                        //����ҵ漼
            pGridColumn[66] = pGrid.GetColumnToIndex("AMOUNT4");                        //�����ֹμ�
            pGridColumn[67] = pGrid.GetColumnToIndex("D27");                        //���������Ư��
            pGridColumn[68] = pGrid.GetColumnToIndex("D29");                        //����ȸ�� //�ĺ����
            pGridColumn[69] = pGrid.GetColumnToIndex("D28");                        //  

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
         
        private int XlLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pPrintingLine, int[] pGridColumn)
        {
            int vXLine = pPrintingLine; //������ ������ ǥ�õǴ� �� ��ȣ

            object vGetValue = null;  

            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;

            string vSUMMARY_FLAG = "N";

            bool IsConvert = false;  
            try
            {
                vSUMMARY_FLAG = iConv.ISNull(pGrid.GetCellValue(pRow, pGridColumn[80]));

                mPrinting.XLActiveSheet(mTargetSheet);

                //[�μ�] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[0]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {
                    
                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {

                }
                else
                {
                    mPrinting.XLCellMerge(vXLine, 2, vXLine + 1, 7, true);                    
                }
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //[���] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[2]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty; 
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine, 5, vConvertString);
                }

                //[����] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[1]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 1, 2, vConvertString);
                }

                //[����] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[3]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 1, 5, vConvertString);
                }
                else
                {
                    mPrinting.XLCellMerge(vXLine + 2, 2, vXLine + 3, 7, true);
                    mPrinting.XLSetCell(vXLine + 2, 2, vConvertString);
                }

                //[�Ի�����] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[5]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 2, 2, vConvertString);
                }                

                //[�޿�����] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[4]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                if (vSUMMARY_FLAG == "N")
                {
                    mPrinting.XLSetCell(vXLine + 2, 5, vConvertString);
                }

                //[�������] 
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[6]);
                IsConvert = IsConvertString(vGetValue, out vConvertString);
                if (IsConvert == true)
                {

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 2, vConvertString);

                //�ٹ����� �� �⺻����.
                //[������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[9]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vConvertDecimal);                    
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 8, vConvertString);

                //[����ٹ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[11]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 11, vConvertString);

                //[�ٹ�(����)]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[73]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //[�̱ٹ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[74]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 17, vConvertString);

                //[����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[75]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vConvertString);

                //[����ð�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[15]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 8, vConvertString);

                //[�߰��ð�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[21]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 11, vConvertString);

                //[���ϱٷ�-��]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[17]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 14, vConvertString);

                //[���ϱٷ�-��]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[19]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.##}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 17, vConvertString);

                //[������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[76]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 20, vConvertString);

                //[����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[77]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 8, vConvertString);

                //[����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[78]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 11, vConvertString);

                //[���°���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[13]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###.###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 14, vConvertString);

                //[���ñ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[10]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 20, vConvertString);
                 
                ////////////////////////////////////////////////////////////////////////////////�����׸�
                //[�⺻��]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[12]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 23, vConvertString);

                ////[�ɾ߼���]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[16]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 1, 23, vConvertString);

                //[��Ÿ����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[14]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 23, vConvertString);

                ////[����������] => ��Ÿ��������.
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[25]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 23, vConvertString);

                //[����������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[23]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 27, vConvertString);

                //[���������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[22]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine , 35, vConvertString);

                //[��Ÿ����2]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[28]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 27, vConvertString);

                //[���ټ���]=> ��Ÿ���翡 ����
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[35]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 27, vConvertString);

                ////[���޼���]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[25]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 31, vConvertString);

                //[���Ͽ���ݾ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[20]);  //35
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 35, vConvertString);

                //[��������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[29]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 31, vConvertString);

                //[���м���
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[31]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 39, vConvertString);

                //[�ޱټ���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[24]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 31, vConvertString);

                //[���ϱٷαݾ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[18]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 31, vConvertString);

                //[�󿩱�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[30]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine , 39, vConvertString);

                ////[��ź�]=>��Ÿ��������
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[30]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 35, vConvertString);

                ////[ö�߼���]=>��Ÿ��������
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[33]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 35, vConvertString);

                //[���Ͽ��� (���� + �ɾ�)]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[27]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine+1, 23, vConvertString);

                //[�ڰݼ���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[26]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 39, vConvertString);

            
                //[�߰��ٹ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[38]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 27, vConvertString);

                //[��������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[36]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 35, vConvertString);

                //[���ܱٷμ���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[37]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 23, vConvertString);

                //[����������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[81]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 31, vConvertString);

                //[���庹����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[82]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 35, vConvertString);

                //[��������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[83]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 39, vConvertString);


                ///////////////////////////////////////////////////////////////////�����׸�//
                //[�ҵ漼]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[41]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 43, vConvertString);

                //[��뺸��]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[44]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 43, vConvertString);

                //[���ڱ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[55]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 43, vConvertString);

                //[���������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[62]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 43, vConvertString);

                //[��������ҵ漼]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[65]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 47, vConvertString);

                //[�ֹμ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[42]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 47, vConvertString);

                //[���ݻ�ȯ]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[46]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 47, vConvertString);

                //[���������]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[50]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 3, 47, vConvertString);

                //[���������ֹμ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[66]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 51, vConvertString);

                //[���ο���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[43]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 51, vConvertString);

                //[���ڳ���]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[49]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 51, vConvertString);

                ////[�ǰ���������]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[47]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 51, vConvertString);

                ////[���������Ư��]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[67]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 51, vConvertString);

                //[�ǰ�����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[45]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 55, vConvertString);

                //[�ĺ����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[68]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 55, vConvertString);
                                

                ////[��纸�������]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[48]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 2, 55, vConvertString);

                //[��Ÿ����]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[54]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 55, vConvertString);

                //�հ�
                //[�����޾�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[40]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 59, vConvertString);

                //[�����Ѿ�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[70]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 1, 59, vConvertString);

                //[�����޾�]
                vGetValue = pGrid.GetCellValue(pRow, pGridColumn[71]);
                IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine + 2, 59, vConvertString);

                ////[]
                //vGetValue = pGrid.GetCellValue(pRow, pGridColumn[67]);
                //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine + 3, 55, vConvertString);

                //���հ� �� �μ� �հ� ���� ����.
                if (vSUMMARY_FLAG == "N")
                {
                    /////////
                }
                else
                {
                    //2.BACK COLOR ����.
                    mPrinting.XLCellColorBrush(mCurrentRow, 2, mCurrentRow + 3, mCopy_EndCol - 1, System.Drawing.Color.LightBlue);
                }

                vXLine = vXLine + 4;
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

        public int XLWirteMain(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid,
                                object pLocal_DATE, object pUserName, object pCorporationName, object pYYYYMM, object pWageTypeName, object pDepartmentName)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;
             
            int[] mGridColumn;

            //�ʱ�ȭ//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 64;
            mCopy_EndRow = 45;

            //mDefaultEndPageRow = 1;
            mDefaultPageRow = 12;    // ������ ������ PageCount �⺻��.
            mPrintingLastRow = 40;  //���� �μ� ����.
            //m1stPrintingLastRow = 40;

            mCurrentRow = 12;
            mCopyLineSUM = 1;

            int vTotalRow = 0;
            int vPageRowCount = 0;  //�μ��� �ش� ���� ���� ����. 

            mCorporationName = pCorporationName;
            mUserName = pUserName;
            mYYYYMM = pYYYYMM;
            mWageTypeName = pWageTypeName;
            mDepartmentName = pDepartmentName;
            mPringingDateTime = pLocal_DATE;

            string vDEPT_CODE = string.Empty;
            object vDEPT_NAME = string.Empty;

            try
            {
                vTotalRow = pGrid.RowCount;
                //TotalPage(pGrid);

                if (vTotalRow > 0)
                {
                    //�迭 ����.
                    SetArray(pGrid, out mGridColumn);
                    vPageRowCount = mCurrentRow - 4;  

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Row : {0} / {1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        if (iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[80])) == "T")
                        {
                            vDEPT_NAME = string.Empty;
                        }
                        else
                        {
                            vDEPT_NAME = pGrid.GetCellValue(vRow, mGridColumn[0]);
                        }

                        if (vRow == 0)
                        {
                            //mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pGrid, vRow, vDEPT_NAME);
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME); 
                        }
                        else if (vDEPT_CODE != iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[79])) && mIsNewPage == false)
                        {
                            XlAllLineClear(pCorporationName);
                            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, vDEPT_NAME);
                            //�����μ� �� �̹Ƿ� ������ROW�� +4�� ����.
                            mCurrentRow = mCurrentRow + (mCopy_EndRow - (vPageRowCount + 4)) + mDefaultPageRow;  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                            vPageRowCount = mDefaultPageRow - 4;
                        }

                        mCurrentRow = XlLine(pGrid, vRow, mCurrentRow, mGridColumn);
                        vPageRowCount = vPageRowCount + 4;
                        if (iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[80])) == "T")
                        {

                        }
                        else
                        {
                            vDEPT_CODE = iConv.ISNull(pGrid.GetCellValue(vRow, mGridColumn[79]));
                        }

                        if (vRow == vTotalRow -1)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            SumWrite(mCurrentRow);      //�հ�.
                            if (vPageRowCount != mPrintingLastRow)
                            {
                                //������ROW�� ������ �μ��ϰ� �ٸ��� ���� ���� CLEAR
                                XlAllLineClear(pCorporationName);
                            }
                        }
                        else
                        {
                            IsNewPage(vPageRowCount, false, vDEPT_NAME);   // ���ο� ������ üũ �� ����.
                            if (mIsNewPage == true)
                            {
                                //�μ� �� �̹Ƿ� ���� ������ROW�� -4�� ����.
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount - 4) + mDefaultPageRow;  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow - 4;
                            }
                        } 
                    }

                    mPrinting.XLDeleteSheet(mSourceSheet1);

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

        private void IsNewPage(int pPrintingLine, bool pIsPageSkep, object pDEPT_NAME)
        {
            if (mPrintingLastRow == pPrintingLine)
            {
                mIsNewPage = true;                
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
            }
            else if (pIsPageSkep == true)
            {
                mIsNewPage = true; 
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, pDEPT_NAME);
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

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, object pDEPT_NAME)
        {
            mPageNumber++; //������ ��ȣ

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(mSourceSheet1); //�� �Լ��� ȣ�� ���� ������ �׸������� XL Sheet�� Insert ���� �ʴ´�.

            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);
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

            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, pDEPT_NAME, mCorporationName);            
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
            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mCorporationName);
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
            mPrinting.XLSetCell((vDrawRow + 0), 59, mCorporationName);

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