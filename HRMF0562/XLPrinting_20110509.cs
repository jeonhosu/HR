using System;
using ISCommonUtil;

namespace HRMF0522
{
    /// <summary>
    /// XLPrint Class�� �̿��� Report�� ���� 
    /// </summary>
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGridAdvEx;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private string mXLOpenFileName = string.Empty;

        private int[] mIndexGridColumns = new int[0] { };

        private int mPositionPrintLineSTART = 1; //���� ��½� ���� ���� �� ��ġ ����
        private int[] mIndexXLWriteColumn = new int[0] { }; //������ ����� �� ��ġ ����

        private int mMaxIncrement = 45; //���� ��µǴ� ���� ���ۺ��� �� ���� ����
        private int mSumPrintingLineCopy = 1; //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ �� ���� �� ��
        private int mMaxIncrementCopy = 70; //�ݺ� ����Ǿ��� ���� �ִ� ����

        private int mXLColumnAreaSTART = 1; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
        private int mXLColumnAreaEND = 45;  //����Ǿ��� ��Ʈ�� ��, ���῭

        #endregion;

        #region ----- Property -----

        /// <summary>
        /// ��� Error Message ���
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        /// <summary>
        /// Message ����� Grid
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISGridAdvEx MessageGridEx
        {
            set
            {
                mGridAdvEx = value;
            }
        }

        /// <summary>
        /// ��ü Data ���� ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar1
        {
            set
            {
                mProgressBar1 = value;
            }
        }

        /// <summary>
        /// Page�� Data ���� ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar2
        {
            set
            {
                mProgressBar2 = value;
            }
        }

        /// <summary>
        /// Ope�� Excel File �̸�
        /// </summary>
        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting()
        {
            mPrinting = new XL.XLPrint();
        }

        #endregion;

        #region ----- Interior Use Methods ----

        #region ----- MessageGrid Methods ----

        private void MessageGrid(string pMessage)
        {
            int vCountRow = mGridAdvEx.RowCount;
            vCountRow = vCountRow + 1;
            mGridAdvEx.RowCount = vCountRow;

            int vCurrentRow = vCountRow - 1;

            mGridAdvEx.SetCellValue(vCurrentRow, 0, pMessage);

            mGridAdvEx.CurrentCellMoveTo(vCurrentRow, 0);
            mGridAdvEx.Focus();
            mGridAdvEx.CurrentCellActivate(vCurrentRow, 0);
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

        #endregion;

        #region ----- XLPrint Define Methods ----

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
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

        #region ----- Line Clear All Methods ----

        private void XlAllLineClear(XL.XLPrint pPrinting)
        {
            int vXLColumn1 = 2;  //No[OPERATION_SEQ_NO]
            int vXLColumn2 = 4;  //������[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //���� ����� �۾� ����[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //���׸���, ���� ��
            int vXLDrawLineColumnEND = 45;  //���׸���, ���� ��

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            //pPrinting.XLActiveSheet(2);
            pPrinting.XLActiveSheet("SourceTab1");

            for (int vXLine = mPositionPrintLineSTART; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //������[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //���� ����� �۾� ����[OPERATION_COMMENT]

                if (vXLine < mMaxIncrementCopy)
                {
                    pPrinting.XL_LineClear(vXLine, vXLDrawLineColumnSTART, vXLDrawLineColumnEND);
                }
            }
        }

        #endregion;

        #region ----- Line Clear Methods ----

        //XlLineClear(mPrinting, vPrintingLine);
        private void XlLineClear(XL.XLPrint pPrinting, int pPrintingLine)
        {
            int vXLColumn1 = 2;  //No[OPERATION_SEQ_NO]
            int vXLColumn2 = 4;  //������[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //���� ����� �۾� ����[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //���׸���, ���� ��
            int vXLDrawLineColumnEND = 45;  //���׸���, ���� ��

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            for (int vXLine = pPrintingLine; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //������[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //���� ����� �۾� ����[OPERATION_COMMENT]

                if (vXLine < mMaxIncrementCopy)
                {
                    pPrinting.XL_LineClear(vXLine, vXLDrawLineColumnSTART, vXLDrawLineColumnEND);
                }
            }
        }

        #endregion;

        #region ----- Define Print Column Methods ----

        private void XLDefinePrintColumn(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            try
            {
                //Grid�� [Edit] ���� [DataColumn] ���� �ִ� �� �̸��� ���� �ϸ� �ȴ�.
                string[] vGridDataColumns = new string[]
                {
                    "NAME",
                    "PERSON_NUM",
                    "DEPT_NAME",
                    "POST_NAME",
                    "JOB_CLASS_NAME",
                    "SUPPLY_DATE",
                    "BANK_NAME",
                    "BANK_ACCOUNTS",
                    "REAL_AMOUNT"
                };

                int vIndexColumn = 0;
                mIndexGridColumns = new int[vGridDataColumns.Length];

                foreach (string vName in vGridDataColumns)
                {
                    mIndexGridColumns[vIndexColumn] = pGrid.GetColumnToIndex(vName);
                    vIndexColumn++;
                }

                //������ ��µ� �� ��ġ ����
                int[] vXLColumns = new int[]
                {
                    28,
                    28,
                    28,
                    29,
                    29,
                    29,
                    30,
                    30,
                    60
                };
                mIndexXLWriteColumn = new int[vXLColumns.Length];
                for (int vCol = 0; vCol < vXLColumns.Length; vCol++)
                {
                    mIndexXLWriteColumn[vCol] = vXLColumns[vCol];
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print HeaderColumns Methods ----

        private void XLHeaderColumns(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory, int pXLine)
        {
            int vXLine = pXLine - 1; //mPositionPrintLineSTART - 1, ��µ� ������ �� ��ġ���� ���� ���� �����Ƿ� 1�� ����.
            int vCountColumn = mIndexGridColumns.Length;

            object vObject = null;
            int vGetIndexGridColumn = 0;

            try
            {
                if (mIndexGridColumns.Length < 1)
                {
                    return;
                }

                //Header Columns
                for (int vCol = 0; vCol < vCountColumn; vCol++)
                {
                    vGetIndexGridColumn = mIndexGridColumns[vCol];
                    switch (pTerritory)
                    {
                        case 1: //Default
                            vObject = pGrid.GridAdvExColElement[vGetIndexGridColumn].HeaderElement[0].Default;
                            mPrinting.XLSetCell(vXLine, mIndexXLWriteColumn[vCol], vObject);
                            break;
                        case 2: //KR
                            vObject = pGrid.GridAdvExColElement[vGetIndexGridColumn].HeaderElement[0].TL1_KR;
                            mPrinting.XLSetCell(vXLine, mIndexXLWriteColumn[vCol], vObject);
                            break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print Content Write Methods ----

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
                        //string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeLong = vDateTime.ToString("yyyy�� MM�� dd��", null);
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

        #endregion

        #region ----- New Page iF Methods ----

        private int NewPage(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTotalRow, int pSumWriteLine)
        {
            int vPrintingRowSTART = 0;
            int vPrintingRowEND = 0;

            try
            {
                vPrintingRowSTART = pSumWriteLine;
                pSumWriteLine = pSumWriteLine + mMaxIncrement;
                vPrintingRowEND = pSumWriteLine;

                //XLContentWrite(mPrinting, pGrid, pTotalRow, mPositionPrintLineSTART, mIndexXLWriteColumn, vPrintingRowSTART, vPrintingRowEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return pSumWriteLine;
        }

        #endregion;

        #region ----- Excel Clear -----

        private void XLContentClear()
        {
            mPrinting.XLActiveSheet("SourceTab1");
    
            // ù�� ���ޱ���/�μ�/����/���/�̸�
            mPrinting.XLSetCell(11, 8, "");   //�μ�
            mPrinting.XLSetCell(13, 8, "");   //����
            mPrinting.XLSetCell(15, 8, "");   //���
            mPrinting.XLSetCell(17, 8, "");   //�̸�                    

            // ��������.
            mPrinting.XLSetCell(26, 9, "");   //����
            mPrinting.XLSetCell(26, 22, "");  //���
            mPrinting.XLSetCell(26, 36, "");  //�μ�
            mPrinting.XLSetCell(27, 9, "");   //����
            mPrinting.XLSetCell(27, 22, "");  //����                    
            mPrinting.XLSetCell(27, 36, "");  //������
            mPrinting.XLSetCell(28, 9, "");   //�Ա�����
            mPrinting.XLSetCell(28, 22, "");  //�Աݰ���

            //============================================================================================
            // �⺻��
            //============================================================================================
            mPrinting.XLSetCell(38, 32, "");
            mPrinting.XLSetCell(38, 36, "");
            mPrinting.XLSetCell(38, 40, "");
            mPrinting.XLSetCell(55, 15, "");
            mPrinting.XLSetCell(54, 34, "");
            mPrinting.XLSetCell(55, 34, "");
            mPrinting.XLSetCell(62, 15, "");
            mPrinting.XLSetCell(61, 34, "");
            mPrinting.XLSetCell(62, 34, "");
            mPrinting.XLSetCell(63, 15, "");
            mPrinting.XLSetCell(63, 34, "");
            mPrinting.XLSetCell(64, 25, "");
            mPrinting.XLSetCell(1, 4, "");
            mPrinting.XLSetCell(67, 4, "");  //���
            
            // ���޿� �����׸�.
            for (int nRow = 0; nRow <= 12; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 6, "");
                mPrinting.XLSetCell(42 + nRow, 15, "");
            }
            // ���޿� �����׸�.
            for (int nRow = 0; nRow <= 11; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 25, "");
                mPrinting.XLSetCell(42 + nRow, 34, "");
            }
            
            //============================================================================================
            // ����(����)
            //============================================================================================
            mPrinting.XLSetCell(32, 12, "");
            mPrinting.XLSetCell(32, 16, "");
            mPrinting.XLSetCell(32, 20, "");
            mPrinting.XLSetCell(33, 8, "");
            mPrinting.XLSetCell(33, 12, "");
            mPrinting.XLSetCell(33, 16, "");
            mPrinting.XLSetCell(34, 8, "");
            mPrinting.XLSetCell(34, 12, "");
            mPrinting.XLSetCell(34, 16, "");
            mPrinting.XLSetCell(38, 4, "");
            mPrinting.XLSetCell(38, 8, "");
            mPrinting.XLSetCell(38, 12, "");
            mPrinting.XLSetCell(38, 16, "");
            mPrinting.XLSetCell(38, 20, "");
            mPrinting.XLSetCell(38, 24, "");
            mPrinting.XLSetCell(38, 28, "");

            // �� �����׸�.
            for (int nRow = 0; nRow <= 5; nRow++)
            {
                mPrinting.XLSetCell(56 + nRow, 6, "");
                mPrinting.XLSetCell(56 + nRow, 15, "");
            }
            // �� �����׸�.
            for (int nRow = 0; nRow <= 4; nRow++)
            {
                mPrinting.XLSetCell(56 + nRow, 25, "");
                mPrinting.XLSetCell(56 + nRow, 34, "");
            }
        }

        private void XLContentClear2()
        {
            mPrinting.XLActiveSheet("SourceTab2");

            // ù�� ���ޱ���/�μ�/����/���/�̸�
            mPrinting.XLSetCell(11, 8, "");   //�μ�
            mPrinting.XLSetCell(13, 8, "");   //����
            mPrinting.XLSetCell(15, 8, "");   //���
            mPrinting.XLSetCell(17, 8, "");   //�̸�                    

            // ��������.
            mPrinting.XLSetCell(26, 9, "");   //����
            mPrinting.XLSetCell(26, 22, "");  //���
            mPrinting.XLSetCell(26, 36, "");  //�μ�
            mPrinting.XLSetCell(27, 9, "");   //����
            mPrinting.XLSetCell(27, 22, "");  //����                    
            mPrinting.XLSetCell(27, 36, "");  //������
            mPrinting.XLSetCell(28, 9, "");   //�Ա�����
            mPrinting.XLSetCell(28, 22, "");  //�Աݰ���

            //============================================================================================
            // �⺻��
            //============================================================================================
            mPrinting.XLSetCell(38, 32, "");
            mPrinting.XLSetCell(38, 36, "");
            mPrinting.XLSetCell(38, 40, "");
            mPrinting.XLSetCell(61, 15, "");
            mPrinting.XLSetCell(61, 34, "");
            mPrinting.XLSetCell(62, 25, "");            
            mPrinting.XLSetCell(1, 4, "");
            mPrinting.XLSetCell(65, 4, "");  //���

            // ���޿� ����/�����׸�.
            for (int nRow = 0; nRow <= 18; nRow++)
            {
                mPrinting.XLSetCell(42 + nRow, 6, "");
                mPrinting.XLSetCell(42 + nRow, 15, "");
                mPrinting.XLSetCell(42 + nRow, 25, "");
                mPrinting.XLSetCell(42 + nRow, 34, "");
            }

            //============================================================================================
            // ����(����)
            //============================================================================================
            mPrinting.XLSetCell(32, 12, "");
            mPrinting.XLSetCell(32, 16, "");
            mPrinting.XLSetCell(32, 20, "");
            mPrinting.XLSetCell(33, 8, "");
            mPrinting.XLSetCell(33, 12, "");
            mPrinting.XLSetCell(33, 16, "");
            mPrinting.XLSetCell(34, 8, "");
            mPrinting.XLSetCell(34, 12, "");
            mPrinting.XLSetCell(34, 16, "");
            mPrinting.XLSetCell(38, 4, "");
            mPrinting.XLSetCell(38, 8, "");
            mPrinting.XLSetCell(38, 12, "");
            mPrinting.XLSetCell(38, 16, "");
            mPrinting.XLSetCell(38, 20, "");
            mPrinting.XLSetCell(38, 24, "");
            mPrinting.XLSetCell(38, 28, "");
        }

        #endregion

        #region ----- XLContent Write -----
		 
        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, int pAllowance_Row, int nAllowance_Column)
        {
            decimal vAMOUNT = 0;
            decimal vDUTY_TIME = 0;
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");
                if (pCnt == 1)
                {                    
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("NAME");                   //����
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("PERSON_NUM");             //���
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DEPT_NAME");              //�μ�
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("POST_NAME");              //����
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("JOB_CLASS_NAME");         //����
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("SUPPLY_DATE");            //������
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("BANK_NAME");              //�Ա�����
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");          //�Աݰ���                 
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("BASIC_AMOUNT");          //�⺻��
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("BASIC_TIME_AMOUNT");     //�ñ�
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("DESCRIPTION");           //���
                    // '���'�� �Ŀ� �߰��� ���Ե� ���̶� Column ������ 15�� �� ����.

                    // ���� Report ��ܿ� ��µ� ����
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT"); //���ñ�
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("WAGE_TYPE");             //�޻󿩱��и�
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("PAY_YYYYMM");            //���޳��                    

                    int vIndexWageType = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //���ޱ���
                    
                    int vIDX_TOT_REAL = pGrid.GetColumnToIndex("REAL_AMOUNT");                // �� �����޾�
                    int vIDX_TOT_SUPP = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          // �����޾�
                    int vIDX_TOT_DED = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");              // �Ѱ�����

                    int vIDX_PAY_REAL = pGrid.GetColumnToIndex("REAL_PAY_AMOUNT");            // �޿� �����޾�
                    int vIDX_PAY_SUPP = pGrid.GetColumnToIndex("TOT_PAY_SUP_AMOUNT");         // �޿� �����޾�
                    int vIDX_PAY_DED = pGrid.GetColumnToIndex("TOT_PAY_DED_AMOUNT");          // �޿� �Ѱ�����

                    int vIDX_BONUS_REAL = pGrid.GetColumnToIndex("REAL_BONUS_AMOUNT");        // �޿� �����޾�
                    int vIDX_BONUS_SUPP = pGrid.GetColumnToIndex("TOT_BONUS_SUP_AMOUNT");     // �޿� �����޾�
                    int vIDX_BONUS_DED = pGrid.GetColumnToIndex("TOT_BONUS_DED_AMOUNT");      // �޿� �Ѱ�����


                    // ù�� ���ޱ���/�μ�/����/���/�̸�
                    mPrinting.XLSetCell(11, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));   //�μ�
                    mPrinting.XLSetCell(13, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //����
                    mPrinting.XLSetCell(15, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));   //���
                    mPrinting.XLSetCell(17, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //�̸�                    

                    // ��������.
                    mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //����
                    mPrinting.XLSetCell(26, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));  //���
                    mPrinting.XLSetCell(26, 36, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));  //�μ�
                    mPrinting.XLSetCell(27, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //����
                    mPrinting.XLSetCell(27, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));  //����                    
                    mPrinting.XLSetCell(27, 36, iString.ISNull(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6)).Substring(0, 10));  //������
                    mPrinting.XLSetCell(28, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));   //�Ա�����
                    mPrinting.XLSetCell(28, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));  //�Աݰ���
                    
                    //============================================================================================
                    // �⺻��
                    //============================================================================================
                    vAMOUNT = 0; 
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 32, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 32, vAMOUNT);
                    }

                    //============================================================================================
                    // �ñ�
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 36, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 36, vAMOUNT);
                    }
                    
                    //============================================================================================
                    // ��� �ñ�
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vAMOUNT  == 0)
                    {
                        mPrinting.XLSetCell(38, 40, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 40, vAMOUNT);
                    }

                    //============================================================================================
                    // �����հ�/�����հ�/�����޾�
                    //============================================================================================
                    // ���޿� 
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_SUPP), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(55, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(55, 15, vAMOUNT);  
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_DED), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(54, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(54, 34, vAMOUNT);   //�Ѱ���
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_PAY_REAL), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(55, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(55, 34, vAMOUNT);  //�����޾�
                    }

                    // �󿩾�
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_SUPP), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(62, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_DED), 0);
                    if (vAMOUNT == 0)  //�Ѱ�����
                    {
                        mPrinting.XLSetCell(61, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 34, vAMOUNT);   //�Ѱ���
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_BONUS_REAL), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(62, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 34, vAMOUNT);  //�����޾�
                    }
                    
                    // �����޾�
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_SUPP), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(63, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(63, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_DED), 0);
                    if (vAMOUNT == 0)  //�Ѱ�����
                    {
                        mPrinting.XLSetCell(63, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(63, 34, vAMOUNT);   //�Ѱ���
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_REAL), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(64, 25, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(64, 25, vAMOUNT);  //�����޾�
                    }

                    mPrinting.XLSetCell(1, 4, pGrid.GetCellValue(pIndexRow, vIndexWageType));
                    mPrinting.XLSetCell(67, 4, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));  //���
                }
                else if (pCnt == 2) {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //���޾�                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(pAllowance_Row+pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row+pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //������                    

                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("OVER_TIME");        //����(����)
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("NIGHT_BONUS_TIME"); //�߰�(����)
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LATE_TIME");        //���°���(����)
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("HOLY_1_TIME");      //�ٹ�(������)
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("HOLY_1_OT");        //����(������)
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("HOLY_1_NIGHT");     //�߰�(������)
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("HOLY_0_TIME");      //�ٹ�(������)
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("HOLY_0_OT");        //����(������)
                    int vIndexDataColumn9 = pGrid.GetColumnToIndex("HOLY_0_NIGHT");     //�߰�(������)
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");   //�ٹ�(�ΰ�����)
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("DUTY_30");         //����(�ΰ�����)
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");  //����(�ΰ�����)
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("HOLY_1_COUNT");    //����(�ΰ�����)
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("HOLY_0_COUNT");    //����(�ΰ�����)
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("TOT_DED_COUNT");   //�̱ٹ�(�ΰ�����)
                    int vIndexDataColumn16 = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");//������(�ΰ�����)

                    //============================================================================================
                    // ����(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 12, "");
                    }
                    else 
                    {
                        mPrinting.XLSetCell(32, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ���°���(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 20, vDUTY_TIME);
                    }                   
                    
                    //============================================================================================
                    // �ٹ�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 12,vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �ٹ�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �ٹ�(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 4, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 4, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �̱ٹ�(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 24, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 24, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ������(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 28, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 28, vDUTY_TIME);
                    }
                }
                else if (pCnt == 5)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //���޾�                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(56 + pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(56 + pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 6)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //������                    

                    mPrinting.XLSetCell(56 + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(56 + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }
        
        private void XLContentWrite2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, int pAllowance_Row, int nAllowance_Column)
        {
            decimal vAMOUNT = 0;
            decimal vDUTY_TIME = 0;
            try
            {
                mPrinting.XLActiveSheet("SourceTab2");
                if (pCnt == 1)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("NAME");                   //����
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("PERSON_NUM");             //���
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DEPT_NAME");              //�μ�
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("POST_NAME");              //����
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("JOB_CLASS_NAME");         //����
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("SUPPLY_DATE");            //������
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("BANK_NAME");              //�Ա�����
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");          //�Աݰ���                 
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("BASIC_AMOUNT");          //�⺻��
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("BASIC_TIME_AMOUNT");     //�ñ�
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("DESCRIPTION");           //���
                    // '���'�� �Ŀ� �߰��� ���Ե� ���̶� Column ������ 15�� �� ����.

                    // ���� Report ��ܿ� ��µ� ����
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT"); //���ñ�
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("WAGE_TYPE");             //�޻󿩱��и�
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("PAY_YYYYMM");            //���޳��                    

                    int vIndexWageType = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //���ޱ���

                    int vIDX_TOT_REAL = pGrid.GetColumnToIndex("REAL_AMOUNT");                // �� �����޾�
                    int vIDX_TOT_SUPP = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");          // �����޾�
                    int vIDX_TOT_DED = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");              // �Ѱ�����

                    // ù�� ���ޱ���/�μ�/����/���/�̸�
                    mPrinting.XLSetCell(11, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));   //�μ�
                    mPrinting.XLSetCell(13, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //����
                    mPrinting.XLSetCell(15, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));   //���
                    mPrinting.XLSetCell(17, 8, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //�̸�                    

                    // ��������.
                    mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));   //����
                    mPrinting.XLSetCell(26, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));  //���
                    mPrinting.XLSetCell(26, 36, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));  //�μ�
                    mPrinting.XLSetCell(27, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));   //����
                    mPrinting.XLSetCell(27, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));  //����                    
                    mPrinting.XLSetCell(27, 36, iString.ISNull(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6)).Substring(0, 10));  //������
                    mPrinting.XLSetCell(28, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));   //�Ա�����
                    mPrinting.XLSetCell(28, 22, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));  //�Աݰ���

                    //============================================================================================
                    // �⺻��
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 32, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 32, vAMOUNT);
                    }

                    //============================================================================================
                    // �ñ�
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 36, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 36, vAMOUNT);
                    }

                    //============================================================================================
                    // ��� �ñ�
                    //============================================================================================
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vAMOUNT == 0)
                    {
                        mPrinting.XLSetCell(38, 40, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 40, vAMOUNT);
                    }

                    //============================================================================================
                    // �����հ�/�����հ�/�����޾�
                    //============================================================================================
                    // �����޾�
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_SUPP), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(61, 15, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 15, vAMOUNT);

                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_DED), 0);
                    if (vAMOUNT == 0)  //�Ѱ�����
                    {
                        mPrinting.XLSetCell(61, 34, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(61, 34, vAMOUNT);   //�Ѱ���
                    }
                    vAMOUNT = 0;
                    vAMOUNT = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIDX_TOT_REAL), 0);
                    if (vAMOUNT == 0)  //�����޾�
                    {
                        mPrinting.XLSetCell(62, 25, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(62, 25, vAMOUNT);  //�����޾�
                    }

                    mPrinting.XLSetCell(1, 4, pGrid.GetCellValue(pIndexRow, vIndexWageType));
                    mPrinting.XLSetCell(67, 4, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));  //���
                }
                else if (pCnt == 2)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ALLOWANCE_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ALLOWANCE_AMOUNT"); //���޾�                    

                    //for (int nRow = pIndexRow; nRow <= (pTotalRow - 1); nRow++)
                    //{
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //}
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("DEDUCTION_NAME");   //�����׸�
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEDUCTION_AMOUNT"); //������                    

                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    mPrinting.XLSetCell(pAllowance_Row + pIndexRow, 34, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("OVER_TIME");        //����(����)
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("NIGHT_BONUS_TIME"); //�߰�(����)
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LATE_TIME");        //���°���(����)
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("HOLY_1_TIME");      //�ٹ�(������)
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("HOLY_1_OT");        //����(������)
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("HOLY_1_NIGHT");     //�߰�(������)
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("HOLY_0_TIME");      //�ٹ�(������)
                    int vIndexDataColumn8 = pGrid.GetColumnToIndex("HOLY_0_OT");        //����(������)
                    int vIndexDataColumn9 = pGrid.GetColumnToIndex("HOLY_0_NIGHT");     //�߰�(������)
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("TOTAL_ATT_DAY");   //�ٹ�(�ΰ�����)
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("DUTY_30");         //����(�ΰ�����)
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("S_HOLY_1_COUNT");  //����(�ΰ�����)
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("HOLY_1_COUNT");    //����(�ΰ�����)
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("HOLY_0_COUNT");    //����(�ΰ�����)
                    int vIndexDataColumn15 = pGrid.GetColumnToIndex("TOT_DED_COUNT");   //�̱ٹ�(�ΰ�����)
                    int vIndexDataColumn16 = pGrid.GetColumnToIndex("WEEKLY_DED_COUNT");//������(�ΰ�����)

                    //============================================================================================
                    // ����(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ���°���(����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(32, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(32, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �ٹ�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(33, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(33, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �ٹ�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �߰�(������)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(34, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(34, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �ٹ�(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 4, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 4, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 8, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 8, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 12, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 12, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 16, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 16, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ����(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 20, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 20, vDUTY_TIME);
                    }

                    //============================================================================================
                    // �̱ٹ�(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 24, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 24, vDUTY_TIME);
                    }

                    //============================================================================================
                    // ������(�ΰ�����)
                    //============================================================================================
                    vDUTY_TIME = 0;
                    vDUTY_TIME = iString.ISDecimaltoZero(pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));
                    if (vDUTY_TIME == 0)
                    {
                        mPrinting.XLSetCell(38, 28, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(38, 28, vDUTY_TIME);
                    }
                }                
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Excel Wirte Methods ----

        // Excel Wirte Methods 1(�޿�/�� �μ�)
        public int XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory
            , string pPeriodFrom, string pUserName, string pCaption, int pCnt)
        {
            string vMessageText = string.Empty;

            int vPageNumber = 0;
            int vTotalRow = pGrid.RowCount; //Grid�� �� ���
            int nAllowance_Row = 42;
            int nAllowance_Column = 6;

            try
            {
                if (pCnt != 1)
                {
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, vTotalRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }
                else if(pCnt == 1)
                {
                    for (int vRow = 0; vRow <= pRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, pRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }

                if (pCnt == 6)
                {
                    //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy, "SourceTab1");
                    XLContentClear();                    
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return vPageNumber;
        }

        // Excel Wirte Methods 2(�޿� �μ�)
        public int XLWirte2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory
            , string pPeriodFrom, string pUserName, string pCaption, int pCnt)
        {
            string vMessageText = string.Empty;

            int vPageNumber = 0;
            int vTotalRow = pGrid.RowCount; //Grid�� �� ���
            int nAllowance_Row = 42;
            int nAllowance_Column = 6;

            try
            {
                if (pCnt != 1)
                {
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite2(pGrid, vRow, vTotalRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }
                else if (pCnt == 1)
                {
                    for (int vRow = 0; vRow <= pRow; vRow++)
                    {
                        vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite2(pGrid, vRow, pRow, pCnt, nAllowance_Row, nAllowance_Column);
                    }
                }

                if (pCnt == 4)
                {
                    //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy, "SourceTab2");
                    XLContentClear2();
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return vPageNumber;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(int pCopySumPrintingLine, string pSourceTab)
        {
            int vPrintHeaderColumnSTART = mXLColumnAreaSTART; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = mXLColumnAreaEND;     //����Ǿ��� ��Ʈ�� ��, ���῭

            int vCopySumPrintingLine = 0;
            vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mMaxIncrementCopy;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                //mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLActiveSheet(pSourceTab); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mMaxIncrementCopy, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

                mPrinting.XLPrinting(1, 1);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return 1; // vCopySumPrintingLine;            
            //mPrinting.XLPrintPreview();
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPrinting(pPageSTART, pPageEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        public void PreView()
        {
            try
            {
                mPrinting.XLPrintPreview();
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            try
            {
                System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
                vMaxNumber = vMaxNumber + 1;
                string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

                vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
                mPrinting.XLSave(vSaveFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

    }
}
#endregion;