using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0240
{
    /// <summary>
    /// XLPrint Class�� �̿��� Report�� ���� 
    /// </summary>
    public class XLPrinting
    {
        #region ----- Variables -----
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGridAdvEx;

        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private string mXLOpenFileName = string.Empty;

        private int[] mIndexGridColumns = new int[0] { };

        private int mPositionPrintLineSTART = 4; //���� ��½� ���� ���� �� ��ġ ����
        private int[] mIndexXLWriteColumn = new int[0] { }; //������ ����� �� ��ġ ����

        //private int mSumWriteLine = 0;      //������ ��µǴ� �� ���� ��
        private int mMaxIncrement = 63;       //���� ��µǴ� ���� ���ۺ��� �� ���� ����
        private int mSumPrintingLineCopy = 1; //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ �� ���� �� ��
        private int mMaxIncrementCopy = 55;   //�ݺ� ����Ǿ��� ���� �ִ� ����

        private int mXLColumnAreaSTART = 1;   //����Ǿ��� ��Ʈ�� ��, ���ۿ�
        private int mXLColumnAreaEND = 40;    //����Ǿ��� ��Ʈ�� ��, ���῭

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

        #region ----- Title Methods ----

        private void XLTitle(int pRow, int pColumn, string pTitle)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLSetCell(pRow, pColumn, pTitle);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Header Left Methods ----

        private void XLHeaderLeft(int pRow, int pColumn, string pContent)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLSetCell(pRow, pColumn, pContent);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Header Right Methods ----

        private void XLHeaderRight(int pRow, int pColumn, string pContent)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLSetCell(pRow, pColumn, pContent);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Footer Left Methods ----

        private void XLFooterLeft(int pRow, int pColumn, string pContent)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLSetCell(pRow, pColumn, pContent);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Footer Right Methods ----

        private void XLFooterRight(int pRow, int pColumn, string pContent)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                mPrinting.XLSetCell(pRow, pColumn, pContent);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print Header Methods ----

        private void XLHeader(string pTitle, string pHeaderLeft, string pHeaderRight)
        {
            try
            {
                XLTitle(6, 14, pTitle);

                //XLHeaderLeft(4, 2, pHeaderLeft);
                //XLHeaderRight(4, 52, pHeaderRight); //����� ������
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Print Footer Methods ----

        private void XLFooter(string pFooterLeft, string pFooterRight)
        {
            try
            {
                XLFooterLeft(44, 2, pFooterLeft);   //�ϴ��� ����
                XLFooterRight(44, 41, pFooterRight);//�ϴ��� ������
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
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
                    "LOAN_NUM",
                    "ISSUE_DATE",
                    "DUE_DATE",
                    "BANK_NAME",
                    "ACCOUNT_DESC",
                    "CURRENCY_CODE",
                    "LOAN_AMOUNT",
                    "LOAN_CURR_AMOUNT",
                    "REPAY_LAST_DATE",
                    "REPAY_COUNT",
                    "REPAY_SUM_AMOUNT",
                    "REPAY_SUM_CURR_AMOUNT",
                    "REPAY_INTEREST_COUNT",
                    "REPAY_INTEREST_SUM_AMOUNT",
                    "REPAY_INTEREST_SUM_CURR_AMOUNT"
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
                    2,  // LOAN_NUM                         ���Թ�ȣ          
                    5,  // ISSUE_DATE                       ��������          
                    9,  // DUE_DATE                         ��������          
                    13, // BANK_NAME                        ��������          
                    17, // ACCOUNT_DESC                     ���԰�����        
                    21, // CURRENCY_CODE                    ��ȭ              
                    25, // LOAN_AMOUNT                      �����ܾ�(��ȭ)    
                    29, // LOAN_CURR_AMOUNT                 �����ܾ�(��ȭ)    
                    33, // REPAY_LAST_DATE                  ������ȯ����      
                    37, // REPAY_COUNT                      ���ݻ�ȯȽ��      
                    41, // REPAY_SUM_AMOUNT                 ��ȯ����(��ȭ)    
                    45, // REPAY_SUM_CURR_AMOUNT            ��ȯ����(��ȭ)    
                    49, // REPAY_INTEREST_COUNT             ���ڻ�ȯȽ��      
                    54, // REPAY_INTEREST_SUM_AMOUNT        ���ڻ�ȯ����(��ȭ)
                    59  // REPAY_INTEREST_SUM_CURR_AMOUNT   ���ڻ�ȯ����(��ȭ)
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
            bool IsConvert = false;

            try
            {
                if (pObject != null)
                {
                    //IsConvert = pObject is System.DateTime;
                    //if (IsConvert == true)
                    //{
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        //string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeLong = vDateTime.ToString("yyyy�� MM�� dd��", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    //}
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

        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow)
        {
            //object vObject = null;

            try
            {
                mPrinting.XLActiveSheet("Sheet1");
                
                int vIndexDataColumn1 = pGrid.GetColumnToIndex("PRINT_NUM");              //�߱޹�ȣ
                int vIndexDataColumn2 = pGrid.GetColumnToIndex("CERTIFICATE_TITLE");    //����
                int vIndexDataColumn3 = pGrid.GetColumnToIndex("NAME");                   //�ѱ�
                int vIndexDataColumn4 = pGrid.GetColumnToIndex("CHINESE_NAME");           //����
                int vIndexDataColumn5 = pGrid.GetColumnToIndex("REPRE_NUM");              //�ֹε�Ϲ�ȣ
                int vIndexDataColumn6 = pGrid.GetColumnToIndex("PERSON_ADDRESS");         //�ּ�
                int vIndexDataColumn7 = pGrid.GetColumnToIndex("DEPT_NAME");              //�μ�
                int vIndexDataColumn8 = pGrid.GetColumnToIndex("POST_NAME");              //����
                int vIndexDataColumn9 = pGrid.GetColumnToIndex("ORI_JOIN_DATE");          //�����Ⱓ(��������)
                int vIndexDataColumn10 = pGrid.GetColumnToIndex("RETIRE_DATE");           //�����Ⱓ(��������)
                int vIndexDataColumn11 = pGrid.GetColumnToIndex("DESCRIPTION");           //�뵵
                int vIndexDataColumn12 = pGrid.GetColumnToIndex("SEND_ORG");              //����ó
                int vIndexDataColumn13 = pGrid.GetColumnToIndex("CERTIFICATE_REMARK");    //Ư�̻���
                int vIndexDataColumn14 = pGrid.GetColumnToIndex("PRINT_COUNT");           //����
                int vIndexDataColumn15 = pGrid.GetColumnToIndex("PRINT_DATE");            //�μ�����
                int vIndexDataColumn16 = pGrid.GetColumnToIndex("CORP_NAME");             //ȸ���
                int vIndexDataColumn17 = pGrid.GetColumnToIndex("CORP_ADDRESS");          //ȸ���ּ�
                int vIndexDataColumn18 = pGrid.GetColumnToIndex("PRESIDENT_NAME");        //��ǥ�̻�
                int vIndexDataColumn19 = pGrid.GetColumnToIndex("WORKING_NAME");          //������

                //�߱޹�ȣ
                mPrinting.XLSetCell(11, 3, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                
                //����
                //Code(01) : ��������, Code(02) : �������, Code(03) : ��������
                object vCertificate_Code = pGrid.GetCellValue(pIndexRow, vIndexDataColumn2);
                mPrinting.XLSetCell(2, 2, vCertificate_Code); 
                
                //�ѱ�
                mPrinting.XLSetCell(14, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                ////����
                //mPrinting.XLSetCell(15, 13, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));

                //�ֹι�ȣ
                mPrinting.XLSetCell(14, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));

                //�ּ�
                mPrinting.XLSetCell(17, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));

                //����
                mPrinting.XLSetCell(20, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));

                //�μ�
                mPrinting.XLSetCell(20, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                  
                //�����Ⱓ(��������)
                //if(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9) != null)
                //{
                //    object test1 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    //mPrinting.XLSetCell(23, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                //}
                //else
                //    mPrinting.XLSetCell(30, 13, "");

                //�����Ⱓ(��������)
                //if (pGrid.GetCellValue(pIndexRow, vIndexDataColumn10) != null)
                //{
                //    object test2 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                mPrinting.XLSetCell(23, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                
                //}
                //else
                //    mPrinting.XLSetCell(33, 13, "");

                //������
                mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn19));

                //�뵵
                mPrinting.XLSetCell(29, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));

                ////����ó
                //mPrinting.XLSetCell(39, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));

                //����
                mPrinting.XLSetCell(33, 3, pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));

                ////�� ��
                //mPrinting.XLSetCell(45, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));

                //�μ�����
                mPrinting.XLSetCell(41, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));                

                //ȸ���
                mPrinting.XLSetCell(43, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));

                //ȸ���ּ�
                mPrinting.XLSetCell(46, 2, pGrid.GetCellValue(pIndexRow, vIndexDataColumn17));

                //��ǥ�ڸ�
                mPrinting.XLSetCell(49, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn18));

            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Excel Wirte Methods ----

        public int XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int nPrintTotalCnt, int pTerritory, string pPeriodFrom, /*string pPeriodTo,*/ string pUserName, string pCaption, string pLang)
        {
            string vMessageText = string.Empty;

            //int vPageNumber = 0;

            try
            {
                //int vTotalRow = pGrid.RowCount; //Grid�� �� ���

                for (int nPrintCnt = 0; nPrintCnt < nPrintTotalCnt; nPrintCnt++)
                {
                    //vPageNumber++;
                    //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy, pLang);

                    //[Content_Printing]
                    XLContentWrite(pGrid, 0);

                  
                }                
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return nPrintTotalCnt;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(int pCopySumPrintingLine, string pLang)
        {
            int vPrintHeaderColumnSTART = mXLColumnAreaSTART; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = mXLColumnAreaEND;     //����Ǿ��� ��Ʈ�� ��, ���῭

            int vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mMaxIncrementCopy;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                if (pLang == "EN")
                {
                    mPrinting.XLActiveSheet("Destination2"); //mPrinting.XLActiveSheet(2);
                }
                else
                {
                    mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(2);
                }
                object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mMaxIncrementCopy, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ

                mPrinting.XLActiveSheet("Sheet1"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vCopySumPrintingLine;
            mPrinting.XLPrintPreview();
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLDeleteSheet("Destination");
                mPrinting.XLDeleteSheet("Destination2");
                mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
                //mPrinting.XLPrinting(pPageSTART, pPageEND);
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