using System;

namespace HRMF0206
{
    /// <summary>
    /// XLPrint Class�� �̿��� Report�� ���� 
    /// </summary>
    public class XLPrinting
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGridAdvEx;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        private int[] mIndexGridColumns = new int[0] { };

        private int mPositionPrintLineSTART = 1; //���� ��½� ���� ���� �� ��ġ ����
        private int[] mIndexXLWriteColumn = new int[0] { }; //������ ����� �� ��ġ ����

        private int mMaxIncrement = 41; //���� ��µǴ� ���� ���ۺ��� �� ���� ����
        private int mSumPrintingLineCopy = 1; //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ �� ���� �� ��
        private int mMaxIncrementCopy = 67; //�ݺ� ����Ǿ��� ���� �ִ� ����

        private int mXLColumnAreaSTART = 1; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
        private int mXLColumnAreaEND = 45;  //����Ǿ��� ��Ʈ�� ��, ���῭

        //���� ����//
        private int mSTART_COL = 1;
        private int mSTART_ROW = 1;
        private int mEND_COL = 45;
        private int mEND_ROW = 41;

        private int mCURRENT_ROW = 0;

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

        #region ----- Report Title -----

        private void ReportTitle()
        {
            //======================================================================================
            // ���� �� �⺻���� �׸�� ��� �κ�
            //======================================================================================
            //����
            mPrinting.XLSetCell(5, 13, "��   ��   ��   ��   ǥ");
            //�⺻����
            mPrinting.XLSetCell(10, 3, "��   ��   ��   ��");
            //����
            mPrinting.XLSetCell(12, 11, "��      ��");
            //����
            mPrinting.XLSetCell(12, 22, "��      ��");
            //�޿�����
            mPrinting.XLSetCell(12, 33, "�޿�����");
            //�μ�
            mPrinting.XLSetCell(14, 11, "��      ��");
            //��å
            mPrinting.XLSetCell(14, 22, "��      å");
            //�ֹε�Ϲ�ȣ
            mPrinting.XLSetCell(14, 33, "�ֹε�Ϲ�ȣ");
            //���
            mPrinting.XLSetCell(16, 11, "��      ��");
            //����
            mPrinting.XLSetCell(16, 22, "��      ��");
            //�������
            mPrinting.XLSetCell(16, 33, "�������");
            //�Ի�����
            mPrinting.XLSetCell(18, 11, "�Ի�����");
            //����
            mPrinting.XLSetCell(18, 22, "��      ��");
            //����
            mPrinting.XLSetCell(18, 33, "��      ��");
            //��ȭ��ȣ
            mPrinting.XLSetCell(20, 11, "��ȭ��ȣ");
            //�̸���
            mPrinting.XLSetCell(20, 22, "�� �� ��");
            //��������
            mPrinting.XLSetCell(20, 33, "��������");
            //======================================================================================
            // å���ӱ� �׸��
            //======================================================================================
            //å���ӱ�
            mPrinting.XLSetCell(44, 25, "å���ӱ�");
            //����Ⱓ
            mPrinting.XLSetCell(44, 27, "����Ⱓ");
            //======================================================================================
            // �з»��� �׸��
            //======================================================================================
            //�з»���
            mPrinting.XLSetCell(23, 3, "�з»���");
            //���
            mPrinting.XLSetCell(23, 5, "�� ��");
            //��ű�
            mPrinting.XLSetCell(23, 10, "�� �� ��");
            //�з�
            mPrinting.XLSetCell(23, 16, "�� ��");
            //����
            mPrinting.XLSetCell(23, 19, "�� ��");
            //======================================================================================
            // �������� �׸��
            //======================================================================================
            //��������
            mPrinting.XLSetCell(23, 25, "��������");
            //����
            mPrinting.XLSetCell(23, 27, "�� ��");
            //����
            mPrinting.XLSetCell(23, 30, "�� ��");
            //�������
            mPrinting.XLSetCell(23, 34, "�������");
            //�з�
            mPrinting.XLSetCell(23, 38, "�� ��");
            //�ٹ�ó
            mPrinting.XLSetCell(23, 41, "�� �� ó");
            //======================================================================================
            // �ڰݻ��� �׸��
            //======================================================================================
            //�ڰ�/����
            mPrinting.XLSetCell(32, 3, "�ڰ�/����");
            //�ڰ�����
            mPrinting.XLSetCell(32, 5, "�ڰ�����");
            //���
            mPrinting.XLSetCell(32, 12, "���");
            //�����
            mPrinting.XLSetCell(32, 17, "�����");
            //======================================================================================
            // ��»��� �׸��
            //======================================================================================
            //��»���
            mPrinting.XLSetCell(32, 25, "��»���");
            //�ٹ��Ⱓ
            mPrinting.XLSetCell(32, 27, "�ٹ��Ⱓ");
            //�ٹ�ó
            mPrinting.XLSetCell(32, 33, "�ٹ�ó");
            //����
            mPrinting.XLSetCell(32, 38, "����");
            //������
            mPrinting.XLSetCell(32, 41, "������");
            //======================================================================================
            // ���л��� �׸��
            //======================================================================================
            //����
            mPrinting.XLSetCell(38, 3, "�� ��");
            //���б���
            mPrinting.XLSetCell(38, 5, "���б���");
            //����
            mPrinting.XLSetCell(38, 11, "�� ��");
            //���
            mPrinting.XLSetCell(38, 17, "���");
            //����
            mPrinting.XLSetCell(38, 20, "�� ��");
            //======================================================================================
            // ǥâ/¡����� �׸��
            //======================================================================================
            //ǥâ/¡��
            mPrinting.XLSetCell(38, 25, "ǥâ/¡��");
            //�������
            mPrinting.XLSetCell(38, 27, "�������");
            //�������
            mPrinting.XLSetCell(38, 33, "�������");
            //����
            mPrinting.XLSetCell(38, 37, "����");
            //����
            mPrinting.XLSetCell(38, 41, "����");
            //======================================================================================
            // �������� �׸��
            //======================================================================================
            //����
            mPrinting.XLSetCell(44, 3, "�� ��");
            //��������
            mPrinting.XLSetCell(44, 5, "��������");
            //�Ⱓ
            mPrinting.XLSetCell(44, 11, "�� ��");
            //������
            mPrinting.XLSetCell(44, 18, "������");
            //======================================================================================
            // �߷ɻ��� �׸��
            //======================================================================================
            //�߷��̷�
            mPrinting.XLSetCell(50, 3, "�� �� �� ��");
            //�߷�����
            mPrinting.XLSetCell(50, 5, "�߷�����");
            //�߷�
            mPrinting.XLSetCell(50, 12, "�� ��");
            //�μ�
            mPrinting.XLSetCell(50, 18, "�� ��");
            //��å
            mPrinting.XLSetCell(50, 23, "�� å");
            //����
            mPrinting.XLSetCell(50, 27, "�� ��");
            //����
            mPrinting.XLSetCell(50, 31, "�� ��");
            //ȣ��
            mPrinting.XLSetCell(50, 35, "ȣ ��");
            //���
            mPrinting.XLSetCell(50, 39, "�� ��");
            //======================================================================================
            // ��ü/���»��� �׸��
            //======================================================================================
            //����
            mPrinting.XLSetCell(28, 3, "����");
            //���
            mPrinting.XLSetCell(28, 16, "���");
            //��ü
            mPrinting.XLSetCell(29, 3, "��ü");
            //����
            mPrinting.XLSetCell(29, 16, "����");
            //�ּ�
            mPrinting.XLSetCell(30, 3, "�ּ�");
            //======================================================================================
            // ���� �ϴ��� ��� ���� �׸��
            //======================================================================================
            //�����
            mPrinting.XLSetCell(65, 27, "����� : ");
            //�������
            mPrinting.XLSetCell(65, 37, "������� : ");
        }

        public void ReportTitle2()
        {
            //======================================================================================
            // ���� �� �⺻���� �׸�� ��� �κ�
            //======================================================================================
            //����
            mPrinting.XLSetCell(1, 12, "[�� �� �� �� ī ��]");

            //�⺻����
            mPrinting.XLSetCell(4, 2, "�� �� �� ��");
            mPrinting.XLSetCell(5, 10, "�� �� ��");
            mPrinting.XLSetCell(5, 28, "�������");
            mPrinting.XLSetCell(6, 10, "��    ��");
            mPrinting.XLSetCell(6, 28, "��    å");
            mPrinting.XLSetCell(7, 10, "��    ��");
            mPrinting.XLSetCell(7, 28, "�޿�����");
            mPrinting.XLSetCell(8, 10, "�������");
            mPrinting.XLSetCell(8, 28, "����(��)");
            mPrinting.XLSetCell(9, 10, "�� �� ��");
            mPrinting.XLSetCell(9, 28, "�� �� ��");
            mPrinting.XLSetCell(10, 10, "�Ի�����");
            mPrinting.XLSetCell(10, 28, "�� �� ó");
            mPrinting.XLSetCell(11, 28, "�� �� ��");

            //======================================================================================
            // �з»���
            //======================================================================================
            mPrinting.XLSetCell(12, 2, "�з»���");

            mPrinting.XLSetCell(13, 2, "���");
            mPrinting.XLSetCell(13, 9, "�б���");
            mPrinting.XLSetCell(13, 15, "��������");
            mPrinting.XLSetCell(13, 18, "������");

            //======================================================================================
            //��»���
            //======================================================================================
            mPrinting.XLSetCell(12, 23, "��»���");

            mPrinting.XLSetCell(13, 23, "���");
            mPrinting.XLSetCell(13, 30, "ȸ���");
            mPrinting.XLSetCell(13, 37, "����");
            mPrinting.XLSetCell(13, 40, "������");

            //======================================================================================
            // �λ���
            //======================================================================================
            mPrinting.XLSetCell(19, 2, "�λ���");

            mPrinting.XLSetCell(20, 2, "�򰡳⵵");
            mPrinting.XLSetCell(20, 7, "�򰡵��");
            mPrinting.XLSetCell(20, 14, "���"); 

            //======================================================================================
            // �߷ɻ���
            //======================================================================================
            mPrinting.XLSetCell(19, 23, "�߷ɻ���");

            mPrinting.XLSetCell(20, 23, "�߷�����");
            mPrinting.XLSetCell(20, 27, "�߷ɻ���");
            mPrinting.XLSetCell(20, 34, "�μ���");
            mPrinting.XLSetCell(20, 40, "����");

            //======================================================================================
            // �ڰ�/����
            //======================================================================================            
            mPrinting.XLSetCell(25, 2, "�ڰ�/����");

            mPrinting.XLSetCell(26, 2, "��Ī");
            mPrinting.XLSetCell(26, 8, "���");
            mPrinting.XLSetCell(26, 13, "����");
            mPrinting.XLSetCell(26, 18, "�������");

            //======================================================================================
            // ��������
            //======================================================================================
            mPrinting.XLSetCell(31, 2, "��������");
            
            mPrinting.XLSetCell(32, 2, "������");
            mPrinting.XLSetCell(38, 9, "�����Ⱓ");
            mPrinting.XLSetCell(38, 16, "����ó");

            //======================================================================================
            // ��Ÿ����
            //======================================================================================
            mPrinting.XLSetCell(36, 2, "��Ÿ����");
            
            mPrinting.XLSetCell(37, 2, "����");
            mPrinting.XLSetCell(37, 6, "����");
            mPrinting.XLSetCell(37, 12, "���");

            mPrinting.XLSetCell(38, 2, "����");
            mPrinting.XLSetCell(39, 2, "���");
            mPrinting.XLSetCell(40, 2, "���ƿ���");

            //======================================================================================
            // �������
            //======================================================================================
            mPrinting.XLSetCell(36, 23, "�������");

            //�������
            mPrinting.XLSetCell(37, 23, "�������");
            //�������
            mPrinting.XLSetCell(37, 27, "�������");
            //����
            mPrinting.XLSetCell(37, 30, "����");
            //����
            mPrinting.XLSetCell(37, 37, "����");
        }

        #endregion;

        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, string pPrintDateTime, string pUserName)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                if (pCnt == 1)
                {   
                    // �⺻ ����1
                    int vIndexDataColumn1  = pGrid.GetColumnToIndex("NAME");            // ����
                    int vIndexDataColumn2  = pGrid.GetColumnToIndex("JOB_CLASS_NAME");  // ����
                    int vIndexDataColumn3  = pGrid.GetColumnToIndex("DEPT_NAME");       // �μ�    
                    int vIndexDataColumn4  = pGrid.GetColumnToIndex("ABIL_NAME");       // ��å    
                    int vIndexDataColumn5  = pGrid.GetColumnToIndex("REPRE_NUM");       // �ֹι�ȣ
                    int vIndexDataColumn6  = pGrid.GetColumnToIndex("PERSON_NUM");      // ���    
                    int vIndexDataColumn7  = pGrid.GetColumnToIndex("D_POST_NAME");       // ����    
                    int vIndexDataColumn8  = pGrid.GetColumnToIndex("RETIRE_DATE");     // �������
                    int vIndexDataColumn9  = pGrid.GetColumnToIndex("JOIN_DATE");       // �Ի�����
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("OCPT_NAME");       // ����
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("PRSN_ADDR1");      // �ּ�1
                    int vIndexDataColumn21 = pGrid.GetColumnToIndex("PRSN_ADDR2");      // �ּ�2
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("EMAIL");           // �̸���
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("TELEPHON_NO");     // ��ȭ��ȣ
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("LABOR_UNION_YN");  // ��������

                    //����
                    mPrinting.XLSetCell(12, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //����
                    mPrinting.XLSetCell(12, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //�μ�
                    mPrinting.XLSetCell(14, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //��å
                    mPrinting.XLSetCell(14, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //�ֹι�ȣ
                    mPrinting.XLSetCell(14, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    //���
                    mPrinting.XLSetCell(16, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    //����
                    mPrinting.XLSetCell(16, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    //�������
                    DateTime dRetireDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));
                    object vRetireDate1 = dRetireDate.ToString("yyyy", null);
                    object vRetireDate2 = dRetireDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    if (vRetireDate1.ToString() == "0001")
                    {
                        mPrinting.XLSetCell(16, 38, "");
                    }
                    else
                    {
                        mPrinting.XLSetCell(16, 38, vRetireDate2);
                    }                  
                    //�Ի�����
                    DateTime dJoinDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    object vJoinDate = dJoinDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(18, 16, vJoinDate);
                    //����
                    mPrinting.XLSetCell(18, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    //�ּ�
                    object vAddress = string.Format("{0} {1}", pGrid.GetCellValue(pIndexRow, vIndexDataColumn11), pGrid.GetCellValue(pIndexRow, vIndexDataColumn21));
                    mPrinting.XLSetCell(30, 5, vAddress);
                    //�̸���
                    mPrinting.XLSetCell(20, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    //��ȭ��ȣ
                    mPrinting.XLSetCell(20, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    //��������
                    object vLaborUnion = pGrid.GetCellValue(pIndexRow, vIndexDataColumn14);
                    if (vLaborUnion.ToString() == "N")
                    {
                        mPrinting.XLSetCell(20, 38, "�̰���");
                    }
                    else if (vLaborUnion.ToString() == "Y")
                    {
                        mPrinting.XLSetCell(20, 38, "����");
                    }
                    else
                    {
                        mPrinting.XLSetCell(20, 38, ""); 
                    }
                }
                else if (pCnt == 2)
                {                 
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("PAYMENT_DATE");   // ����Ⱓ
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("BANK_NAME");      // �����  
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");  // ���¹�ȣ
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("PAY_TYPE_NAME");  // �޿�����

                    //����Ⱓ
                    mPrinting.XLSetCell(45 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //�����
                    mPrinting.XLSetCell(18, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //���¹�ȣ
                    mPrinting.XLSetCell(19, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //�޿�����
                    mPrinting.XLSetCell(12, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("SCHOLARSHIP_TYPE_NAME"); // �з�         
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("GRADUATION_YYYYMM");     // ��������
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("SCHOOL_NAME");           // ��ű�
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("SPECIAL_STUDY_NAME");    // ����                

                    //�з�
                    mPrinting.XLSetCell(24 + pIndexRow, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //��������
                    mPrinting.XLSetCell(24 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //��ű�
                    mPrinting.XLSetCell(24 + pIndexRow, 10, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //���� 
                    mPrinting.XLSetCell(24 + pIndexRow, 19, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("FAMILY_NAME");    // ����    
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("RELATION_NAME");  // ����    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BIRTHDAY");       // �������
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("COMPANY_NAME");   // ȸ��� 
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("END_SCH_NAME");   // �з�

                    //���� 
                    mPrinting.XLSetCell(24 + pIndexRow, 30, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //����
                    mPrinting.XLSetCell(24 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //������� 
                    DateTime vBirthday = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    string sBirthday = vBirthday.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(24 + pIndexRow, 34, sBirthday);
                    //ȸ���
                    mPrinting.XLSetCell(24 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //�з�
                    mPrinting.XLSetCell(24 + pIndexRow, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                }
                else if (pCnt == 5)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("LICENSE_NAME");         // �ڰ�����
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("LICENSE_GRADE_NAME");   // �ڰݵ��
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LICENSE_DATE");         // �������

                    //�ڰ�����
                    mPrinting.XLSetCell(33 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //�ڰݵ��
                    mPrinting.XLSetCell(33 + pIndexRow, 12, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //�������
                    mPrinting.XLSetCell(33 + pIndexRow, 17, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                }
                else if (pCnt == 6)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("COMPANY_NAME");   // �ٹ�ó  
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("POST_NAME");      // ����    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("JOB_NAME");       // ������
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("START_DATE");     // �Ի���  
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("END_DATE");       // �����  

                    //�ٹ�ó
                    mPrinting.XLSetCell(33 + pIndexRow, 33, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //����
                    mPrinting.XLSetCell(33 + pIndexRow, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //������
                    mPrinting.XLSetCell(33 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                    //�Ի���                   
                    DateTime dStartDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow,vIndexDataColumn4));
                    string sStartDate = dStartDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    //�����
                    DateTime dEndDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    string sEndDate = dEndDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);

                    object vStartEndDate = sStartDate + " ~ " + sEndDate;
                    mPrinting.XLSetCell(33 + pIndexRow, 27, vStartEndDate);
                }
                else if (pCnt == 7)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("LANGUAGE_NAME");  // ���б���
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("EXAM_NAME");      // ��������
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("EXAM_LEVEL");     // ���    
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("SCORE");          // ����    

                    //���б���
                    mPrinting.XLSetCell(39 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //��������
                    mPrinting.XLSetCell(39 + pIndexRow, 11, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //���
                    mPrinting.XLSetCell(39 + pIndexRow, 17, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //����
                    mPrinting.XLSetCell(39 + pIndexRow, 20, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 8)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("RP_TYPE_NAME");    // �������
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("RP_NAME");         // �������
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("RP_DATE");         // �������
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("RP_DESCRIPTION");  // �������

                    //�������
                    mPrinting.XLSetCell(39 + pIndexRow, 33, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //�������
                    mPrinting.XLSetCell(39 + pIndexRow, 37, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //�������
                    DateTime dRP_Date = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    string sRP_Date = dRP_Date.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(39 + pIndexRow, 27, sRP_Date);

                    //�������
                    mPrinting.XLSetCell(39 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 9)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("START_DATE");      // ��������
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("END_DATE");        // ��������
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("EDU_ORG");         // ��������
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("EDU_CURRICULUM");  // ��������

                    //��������
                    DateTime dStartDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    string sStartDate = dStartDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    //��������
                    DateTime dEndDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    string sEndDate = dEndDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);

                    object vStartEndDate = sStartDate + " ~ " + sEndDate;
                                        
                    //��������~��������
                    mPrinting.XLSetCell(45 + pIndexRow, 11, vStartEndDate);
                    //��������
                    mPrinting.XLSetCell(45 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //��������
                    mPrinting.XLSetCell(45 + pIndexRow, 18, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 10)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("CHARGE_DATE");    // �߷�����
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("CHARGE_NAME");    // �߷�    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DESCRIPTION");    // ���    
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("DEPT_NAME");      // �μ�    
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("POST_NAME");      // ����    
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("ABIL_NAME");      // ��å    
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("PAY_GRADE_NAME"); // ����    

                    //�߷�����
                    DateTime dChargeDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    object vChargeDate = dChargeDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(51 + pIndexRow, 5, vChargeDate);
                    //�߷�
                    mPrinting.XLSetCell(51 + pIndexRow, 12, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //���
                    mPrinting.XLSetCell(51 + pIndexRow, 39, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //�μ�
                    mPrinting.XLSetCell(51 + pIndexRow, 18, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //����
                    mPrinting.XLSetCell(51 + pIndexRow, 31, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    //��å
                    mPrinting.XLSetCell(51 + pIndexRow, 23, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    //����
                    mPrinting.XLSetCell(51 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                }
                else if (pCnt == 11)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ARMY_KIND_NAME");     // ����
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ARMY_GRADE_NAME");    // ��� 
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("ARMY_END_TYPE_NAME"); // ��������
                    //int vIndexDataColumn4 = pGrid.GetColumnToIndex("DESCRIPTION");      // ����

                    // �����׸� - ����, ���, ��������
                    object vArmyInfo = pGrid.GetCellValue(pIndexRow, vIndexDataColumn1).ToString() + ", "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn2).ToString() + ", "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn3).ToString();

                    mPrinting.XLSetCell(29 + pIndexRow, 19, vArmyInfo);

                    //����
                    //mPrinting.XLSetCell(28 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 12)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("HEIGHT");         // Ű    
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("WEIGHT");         // ������
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BLOOD_NAME");     // ������
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("DISABLED_NAME");  // ���
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("DESCRIPTION");    // ����
                    
                    // ��ü�׸� - Ű, ������, ������
                    object vBodyInfo = pGrid.GetCellValue(pIndexRow, vIndexDataColumn1).ToString() + "cm, "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn2).ToString() + "kg, "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn3).ToString();

                    mPrinting.XLSetCell(29 + pIndexRow, 5, vBodyInfo);

                    //���
                    mPrinting.XLSetCell(28 + pIndexRow, 19, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //����
                    mPrinting.XLSetCell(28 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Excel Open and Close ----

        public void XLOpenClose()
        {
            mPrinting.XLOpenFileClose();

            XLFileOpen();
        }
        #endregion;

        #region ----- Excel Wirte Methods ----

        public void XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory, string pPrintDateTime, string pUserName, string pImageName, int pCnt)
        {
            string vMessageText = string.Empty;

            //int vPageNumber = 0;
            int vTotalRow = pGrid.RowCount; // Grid�� �� ���

            try
            {               
                if (pCnt == 1)
                {
                    for (int vRow = 0; vRow <= pRow; vRow++)
                    {
                        //vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, pRow, pCnt, pPrintDateTime, pUserName);
                    }
                }
 
                if (pCnt != 1)
                {
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        //vPageNumber++;

                        //[Content_Printing]
                        XLContentWrite(pGrid, vRow, vTotalRow, pCnt, pPrintDateTime, pUserName);
                    }
                }

                if (pCnt == 12) // 12��° ������ Grid�� ���,
                {
                    //----------------------------------------[ ������� ��� �κ� ]------------------------------------------
                    if (pRow != 0)
                    {
                        int vIndexImage = mPrinting.CountBarCodeImage;
                        int vCountImage = mPrinting.CountBarCodeImage;
                        for (int vRow = 0; vRow < vCountImage; vRow++)
                        {
                            mPrinting.XLDeleteBarCode(vIndexImage);
                            vIndexImage--;
                        }

                        mPrinting.CountBarCodeImage = 0;
                    }

                    System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                    System.Drawing.PointF vPoint = new System.Drawing.PointF(25F, 125F);
                    mPrinting.XLBarCode(pImageName, vSize, vPoint);
                    //--------------------------------------------------------------------------------------------------------

                    //�λ系�� ������ �׸���� ������ִ� �Լ� ȣ��
                    ReportTitle();

                    //���� �ϴܿ� ��� ���� ǥ��
                    mPrinting.XLSetCell(65, 31, pUserName);
                    mPrinting.XLSetCell(65, 41, pPrintDateTime);

                    //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy);

                    //-------------------------------------------------------------------------------------------------------
                    // ������ ���� ���� �κ�
                    // (SourceTab1�� ������ ��� -> Destination�� ���� -> SourceTab1 ������ ���� ��, ���� ������ ��� 
                    //-------------------------------------------------------------------------------------------------------
                    mPrinting.XLActiveSheet("SourceTab1");
                    int vStartRow = mPositionPrintLineSTART; //���� �� ��ġ ����
                    int vStartCol = mXLColumnAreaSTART;  // +1
                    int vEndRow = mMaxIncrementCopy; // -2
                    int vEndCol = mXLColumnAreaEND;  // -1
                    mPrinting.XLSetCell(vStartRow, vStartCol, vEndRow, vEndCol, null);                  
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        public int XLWirte_PERSON(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, string pPrintDateTime, string pUserName, string pImageName)
        {
            int vRow = 5;

            mPrinting.XLActiveSheet("SourceTab1");

            try
            {
                //REMARK
                int vIDX_Col0 = pGrid.GetColumnToIndex("REMARK");               // ����
                int vIDX_Col0_1 = pGrid.GetColumnToIndex("PERSON_NUM");         // �����ȣ

                // �⺻ ����1
                int vIDX_Col1 = pGrid.GetColumnToIndex("DEPT_NAME");            // �μ�
                int vIDX_Col2 = pGrid.GetColumnToIndex("CONTRACT_TYPE_NAME");   // �������
                int vIDX_Col3 = pGrid.GetColumnToIndex("D_POST_NAME");            // ���� 
                int vIDX_Col4 = pGrid.GetColumnToIndex("ABIL_NAME");            // ��å    
                int vIDX_Col5 = pGrid.GetColumnToIndex("NAME");                 // ����
                int vIDX_Col6 = pGrid.GetColumnToIndex("PAY_TYPE_NAME");        // �޿�����
                int vIDX_Col7 = pGrid.GetColumnToIndex("BIRTHDAY");             // �������
                int vIDX_Col8 = pGrid.GetColumnToIndex("REAL_AGE");             // ������
                int vIDX_Col9 = pGrid.GetColumnToIndex("JOIN_DATE");            // �Ի�����
                int vIDX_Col10 = pGrid.GetColumnToIndex("RETIRE_DATE");         // �������
                int vIDX_Col11 = pGrid.GetColumnToIndex("JOIN_NAME");           // �Ի�����
                int vIDX_Col12 = pGrid.GetColumnToIndex("HP_PHONE_NO");         // ����ó
                int vIDX_Col13 = pGrid.GetColumnToIndex("PRSN_ADDR");           // ���ּ�

                int vIDX_Col14 = pGrid.GetColumnToIndex("ARMY_END_TYPE_NAME");  // ��������
                int vIDX_Col15 = pGrid.GetColumnToIndex("ARMY_PERIOD");         // �����Ⱓ
                int vIDX_Col16 = pGrid.GetColumnToIndex("DISABLED_NAME");       // ��ֳ���
                int vIDX_Col17 = pGrid.GetColumnToIndex("DISABLED_TYPE");       // ��ֵ��
                int vIDX_Col18 = pGrid.GetColumnToIndex("BOHUN_NAME");          // ���ƴ��
                int vIDX_Col19 = pGrid.GetColumnToIndex("BOHUN_RELATION_NAME"); // ���ư���

                //����
                mPrinting.XLSetCell(3, 35, pGrid.GetCellValue(pRow, vIDX_Col0));
                mPrinting.XLSetCell(4, 35, string.Format("[{0}]", pGrid.GetCellValue(pRow, vIDX_Col0_1)));
 
                //�μ�
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col1));
                //�������
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col2));

                //--//
                vRow++;

                //����
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col3));
                //��å
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col4));

                //--//
                vRow++;

                //����
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col5));
                //�޿�����
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col6));

                //--//
                vRow++;

                //�������
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col7));
                //������
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col8));

                //--//
                vRow++;

                //�Ի�����
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col9));
                //�������
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col10));

                //--//
                vRow++;

                //�Ի�����
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col11));
                //����ó
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col12));

                //--//
                vRow++;

                //���ּ�
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col13));

                //��Ÿ����
                vRow = 38;
                //���� ����
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col14));
                //���� ���
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col15));

                //--//
                vRow++;

                //��� ����
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col16));
                //��� ���
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col17));

                //--//
                vRow++;

                //���� ����
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col18));
                //���� ���
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col19));
            }
            catch
            {
                return 1;
            }

            //----------------------------------------[ ������� ��� �κ� ]------------------------------------------
            if (pRow != 0)
            {
                try
                {
                    int vIndexImage = mPrinting.CountBarCodeImage;
                    int vCountImage = mPrinting.CountBarCodeImage;
                    for (int r = 0; r < vCountImage; r++)
                    {
                        mPrinting.XLDeleteBarCode(vIndexImage);
                        vIndexImage--;
                    }
                    mPrinting.CountBarCodeImage = 0;
                }
                catch
                {
                    return 1;
                }
            }

            try
            {
                System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 124.99701F);
                System.Drawing.PointF vPoint = new System.Drawing.PointF(13F, 73F);
                mPrinting.XLBarCode(pImageName, vSize, vPoint);
                //--------------------------------------------------------------------------------------------------------

                //���� �ϴܿ� ��� ���� ǥ��
                mPrinting.XLSetCell(41, 1, pUserName);
                mPrinting.XLSetCell(41, 30, pPrintDateTime);
            }
            catch
            {
                return 1;
            }

            //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
            mSumPrintingLineCopy = CopyAndPaste2(mSumPrintingLineCopy);

            //-------------------------------------------------------------------------------------------------------
            // ������ ���� ���� �κ�
            // (SourceTab1�� ������ ��� -> Destination�� ���� -> SourceTab1 ������ ���� ��, ���� ������ ��� 
            //-------------------------------------------------------------------------------------------------------
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");
                //���� �ʱ�ȭ//           
                mPrinting.XLSetCell(mSTART_ROW, mSTART_COL, mEND_ROW, mEND_COL, null);
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_SCHOLARSHIP(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //�з»��� 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 14;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("SCHOLARSHIP_PERIOD"); // ���         
                int vIDX_Col2 = pGrid.GetColumnToIndex("SCHOOL_NAME");     // �б���
                int vIDX_Col3 = pGrid.GetColumnToIndex("GRADUATION_TYPE_NAME");           // ��������
                int vIDX_Col4 = pGrid.GetColumnToIndex("ADDRESS");    // ������            
                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //�з�
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    //��������
                    mPrinting.XLSetCell(vROW + r, 9, pGrid.GetCellValue(r, vIDX_Col2));
                    //��ű�
                    mPrinting.XLSetCell(vROW + r, 15, pGrid.GetCellValue(r, vIDX_Col3));
                    //���� 
                    mPrinting.XLSetCell(vROW + r, 18, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_CAREER(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //��»��� 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 14;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("CAREER_PERIOD");        // ���         
                int vIDX_Col2 = pGrid.GetColumnToIndex("COMPANY_NAME");         // ȸ���
                int vIDX_Col3 = pGrid.GetColumnToIndex("POST_NAME");            // ����
                int vIDX_Col4 = pGrid.GetColumnToIndex("JOB_NAME");             // ������

                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //���
                    mPrinting.XLSetCell(vROW + r, 23, pGrid.GetCellValue(r, vIDX_Col1));
                    //ȸ���
                    mPrinting.XLSetCell(vROW + r, 30, pGrid.GetCellValue(r, vIDX_Col2));
                    //����
                    mPrinting.XLSetCell(vROW + r, 37, pGrid.GetCellValue(r, vIDX_Col3));
                    //������ 
                    mPrinting.XLSetCell(vROW + r, 40, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_RESULT(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //�λ��� 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 21;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("RESULT_YYYY");          // �򰡳⵵         
                int vIDX_Col2 = pGrid.GetColumnToIndex("RES_LVEL");             // �򰡵��
                //int vIDX_Col3 = pGrid.GetColumnToIndex("RES_SCORE");          // ���
                int vIDX_Col4 = pGrid.GetColumnToIndex("DESCRIPTION");          // ���
              
                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //���
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    //ȸ���
                    mPrinting.XLSetCell(vROW + r, 7, pGrid.GetCellValue(r, vIDX_Col2));
                    //����
                    //mPrinting.XLSetCell(vROW + r, 37, pGrid.GetCellValue(r, vIDX_Col3));
                    //������ 
                    mPrinting.XLSetCell(vROW + r, 14, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_PERSON_HISTORY(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //�λ�߷�
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 21;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("CHARGE_DATE");          // �߷�����         
                int vIDX_Col2 = pGrid.GetColumnToIndex("CHARGE_NAME");          // �߷ɸ�Ī
                int vIDX_Col3 = pGrid.GetColumnToIndex("DEPT_NAME");            // �μ�
                int vIDX_Col4 = pGrid.GetColumnToIndex("POST_NAME");            // ����

                for (int r = 0; r < pGrid.RowCount; r++)
                {                    
                    mPrinting.XLSetCell(vROW + r, 23, pGrid.GetCellValue(r, vIDX_Col1));
                    mPrinting.XLSetCell(vROW + r, 27, pGrid.GetCellValue(r, vIDX_Col2));
                    mPrinting.XLSetCell(vROW + r, 34, pGrid.GetCellValue(r, vIDX_Col3));
                    mPrinting.XLSetCell(vROW + r, 40, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_LICENSE(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //�ڰ�/���� 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 27;  //�μ� ��ġ 
            object vOBJECT;
            string vSTRING;
            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("LICENSE_NAME");         // ��Ī         
                int vIDX_Col2 = pGrid.GetColumnToIndex("LICENSE_GRADE_NAME");   // ���
                int vIDX_Col3 = pGrid.GetColumnToIndex("LICENSE_SCORE");        // ����
                int vIDX_Col4 = pGrid.GetColumnToIndex("LICENSE_DATE");         // �������

                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    mPrinting.XLSetCell(vROW + r, 8, pGrid.GetCellValue(r, vIDX_Col2));

                    vOBJECT = pGrid.GetCellValue(r, vIDX_Col3);
                    vSTRING = string.Format("{0:###,###}", vOBJECT);
                    mPrinting.XLSetCell(vROW + r, 13, vSTRING);
                    mPrinting.XLSetCell(vROW + r, 18, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_EDUCATION(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //��������
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 33;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("EDU_CURRICULUM");       // ������         
                int vIDX_Col2 = pGrid.GetColumnToIndex("EDUCATION_PERIOD");     // �����Ⱓ
                int vIDX_Col3 = pGrid.GetColumnToIndex("EDU_ORG");              // ����ó
                //int vIDX_Col4 = pGrid.GetColumnToIndex("LICENSE_DATE");         // �������

                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    mPrinting.XLSetCell(vROW + r, 9, pGrid.GetCellValue(r, vIDX_Col2));
                    mPrinting.XLSetCell(vROW + r, 16, pGrid.GetCellValue(r, vIDX_Col3));
                    //mPrinting.XLSetCell(vROW + r, 18, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        public int XLWirte_REWARD_PUNISHMENT(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            //������� 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 38;  //�μ� ��ġ 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("RP_DATE");          // �������
                int vIDX_Col2 = pGrid.GetColumnToIndex("RP_TYPE_NAME");     // �������
                int vIDX_Col3 = pGrid.GetColumnToIndex("RP_NAME");          // ����
                int vIDX_Col4 = pGrid.GetColumnToIndex("RP_DESCRIPTION");   // ����

                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    mPrinting.XLSetCell(vROW + r, 23, pGrid.GetCellValue(r, vIDX_Col1));
                    mPrinting.XLSetCell(vROW + r, 27, pGrid.GetCellValue(r, vIDX_Col2));
                    mPrinting.XLSetCell(vROW + r, 30, pGrid.GetCellValue(r, vIDX_Col3));
                    mPrinting.XLSetCell(vROW + r, 37, pGrid.GetCellValue(r, vIDX_Col4));
                }
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(int pCopySumPrintingLine)
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

                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mMaxIncrementCopy, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vCopySumPrintingLine;
            //mPrinting.XLPrintPreview();
        }

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste2(int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = mSTART_COL; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = mEND_COL;     //����Ǿ��� ��Ʈ�� ��, ���῭

            int vCopySumPrintingLine = 0;
            vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mEND_ROW;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(mSTART_ROW, mSTART_COL,mEND_ROW, mEND_COL); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vCopySumPrintingLine;
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

                vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder, vSaveFileName);
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