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
    /// XLPrint Class를 이용해 Report물 제어 
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

        private string m_SheetSource1 = "Destination";
        private string m_SheetSource2 = "Destination2";
        private string m_SheetPrint = "Sheet1";

        private int m_Copy_StartCol = 1;
        private int m_Copy_StartRow = 1;
        private int m_Copy_EndCol = 40;
        private int m_Copy_EndRow = 53;

        private int m_History_Row = 37;
        private int m_Current_Row = 0;

        private int m_PageNumber = 0;

        private int[] mIndexGridColumns = new int[0] { };

        private int mPositionPrintLineSTART = 4; //내용 출력시 엑셀 시작 행 위치 지정
        private int[] mIndexXLWriteColumn = new int[0] { }; //엑셀에 출력할 열 위치 지정

        //private int mSumWriteLine = 0;      //엑셀에 출력되는 행 누적 값
        private int mMaxIncrement = 63;       //실제 출력되는 행의 시작부터 끝 행의 범위
        private int mSumPrintingLineCopy = 1; //엑셀의 선택된 쉬트에 복사되어질 시작 행 위치 및 누적 행 값
        private int mMaxIncrementCopy = 55;   //반복 복사되어질 행의 최대 범위

        private int mXLColumnAreaSTART = 1;   //복사되어질 쉬트의 폭, 시작열
        private int mXLColumnAreaEND = 40;    //복사되어질 쉬트의 폭, 종료열

        #endregion;

        #region ----- Property -----

        /// <summary>
        /// 모든 Error Message 출력
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        /// <summary>
        /// Message 출력할 Grid
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISGridAdvEx MessageGridEx
        {
            set
            {
                mGridAdvEx = value;
            }
        }

        /// <summary>
        /// 전체 Data 진행 ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar1
        {
            set
            {
                mProgressBar1 = value;
            }
        }

        /// <summary>
        /// Page당 Data 진행 ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar2
        {
            set
            {
                mProgressBar2 = value;
            }
        }

        /// <summary>
        /// Ope할 Excel File 이름
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
            try
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
            catch
            {

            }
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
            int vXLColumn2 = 4;  //공정명[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //공정 진행시 작업 조건[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //선그리기, 시작 열
            int vXLDrawLineColumnEND = 45;  //선그리기, 종료 열

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            //pPrinting.XLActiveSheet(2);
            pPrinting.XLActiveSheet("SourceTab1");

            for (int vXLine = mPositionPrintLineSTART; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //공정명[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //공정 진행시 작업 조건[OPERATION_COMMENT]

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
            int vXLColumn2 = 4;  //공정명[OPERATION_DESCRIPTION]
            int vXLColumn3 = 11; //공정 진행시 작업 조건[OPERATION_COMMENT]

            int vXLDrawLineColumnSTART = 2; //선그리기, 시작 열
            int vXLDrawLineColumnEND = 45;  //선그리기, 종료 열

            object vObject = null;
            int vMaxPrintingLine = mMaxIncrementCopy;

            for (int vXLine = pPrintingLine; vXLine < vMaxPrintingLine; vXLine++)
            {
                pPrinting.XLSetCell(vXLine, vXLColumn1, vObject); //No[OPERATION_SEQ_NO]
                pPrinting.XLSetCell(vXLine, vXLColumn2, vObject); //공정명[OPERATION_DESCRIPTION]
                pPrinting.XLSetCell(vXLine, vXLColumn3, vObject); //공정 진행시 작업 조건[OPERATION_COMMENT]

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
                //XLHeaderRight(4, 52, pHeaderRight); //상단의 오른쪽
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
                XLFooterLeft(44, 2, pFooterLeft);   //하단의 왼쪽
                XLFooterRight(44, 41, pFooterRight);//하단의 오른쪽
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
                //Grid의 [Edit] 상의 [DataColumn] 열에 있는 열 이름을 지정 하면 된다.
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

                //엑셀에 출력될 열 위치 지정
                int[] vXLColumns = new int[]
                {
                    2,  // LOAN_NUM                         차입번호          
                    5,  // ISSUE_DATE                       차입일자          
                    9,  // DUE_DATE                         만기일자          
                    13, // BANK_NAME                        차입은행          
                    17, // ACCOUNT_DESC                     차입계정명        
                    21, // CURRENCY_CODE                    통화              
                    25, // LOAN_AMOUNT                      차입잔액(원화)    
                    29, // LOAN_CURR_AMOUNT                 차입잔액(외화)    
                    33, // REPAY_LAST_DATE                  최종상환일자      
                    37, // REPAY_COUNT                      원금상환횟수      
                    41, // REPAY_SUM_AMOUNT                 상환누계(원화)    
                    45, // REPAY_SUM_CURR_AMOUNT            상환누계(외화)    
                    49, // REPAY_INTEREST_COUNT             이자상환횟수      
                    54, // REPAY_INTEREST_SUM_AMOUNT        이자상환누계(원화)
                    59  // REPAY_INTEREST_SUM_CURR_AMOUNT   이자상환누계(외화)
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
            int vXLine = pXLine - 1; //mPositionPrintLineSTART - 1, 출력될 내용의 행 위치에서 한행 위에 있으므로 1을 뺀다.
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
                        string vTextDateTimeLong = vDateTime.ToString("yyyy년 MM월 dd일", null);
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


        private void XLContentWrite(string pActiveSheet, InfoSummit.Win.ControlAdv.ISDataAdapter pCert)
        { 
            try
            {
                mPrinting.XLActiveSheet(pActiveSheet);

                //발급번호
                mPrinting.XLSetCell(11, 3, pCert.CurrentRow["PRINT_NUM"]);

                //증명서
                //Code(01) : 재직증명서, Code(02) : 경력증명서, Code(03) : 퇴직증명서 
                mPrinting.XLSetCell(2, 2, pCert.CurrentRow["CERTIFICATE_TITLE"]);

                //한글
                mPrinting.XLSetCell(14, 9, pCert.CurrentRow["NAME"]);

                ////한자
                //mPrinting.XLSetCell(15, 13, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));

                //주민번호
                mPrinting.XLSetCell(14, 27, pCert.CurrentRow["REPRE_NUM"]);

                //주소
                mPrinting.XLSetCell(17, 9, pCert.CurrentRow["NPERSON_ADDRESSAME"]);

                //직위
                mPrinting.XLSetCell(20, 9, pCert.CurrentRow["POST_NAME"]);

                //부서
                mPrinting.XLSetCell(20, 27, pCert.CurrentRow["DEPT_NAME"]);

                //재직기간(최종일자) 
                mPrinting.XLSetCell(23, 9, pCert.CurrentRow["RETIRE_DATE"]);

                //담당업무
                mPrinting.XLSetCell(26, 9, pCert.CurrentRow["TASK_DESC"]);

                //용도
                mPrinting.XLSetCell(29, 9, pCert.CurrentRow["REMARK"]);

                ////제출처
                //mPrinting.XLSetCell(39, 9, pCert.CurrentRow["SEND_ORG"]);

                //증명서 설명
                mPrinting.XLSetCell(33, 3, pCert.CurrentRow["CERTIFICATE_REMARK"]);

                //인쇄일자
                mPrinting.XLSetCell(41, 6, pCert.CurrentRow["PRINT_DATE"]);

                //회사명
                mPrinting.XLSetCell(43, 6, pCert.CurrentRow["CORP_NAME"]);

                //회사주소
                mPrinting.XLSetCell(46, 2, pCert.CurrentRow["CORP_ADDRESS"]);

                //대표자명
                mPrinting.XLSetCell(49, 6, pCert.CurrentRow["PRESIDENT_NAME"]);

                //회사 영문(상단)
                mPrinting.XLSetCell(1, 2, pCert.CurrentRow["CORP_NAME_ENG"]); 
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        private void XLContentWrite(string pActiveSheet, InfoSummit.Win.ControlAdv.ISDataAdapter pCert
                                    , string pREPRE_FLAG, string pHISTORY_FLAG, InfoSummit.Win.ControlAdv.ISDataAdapter pHistory)
        {
            //object vObject = null;

            try
            { 
                mPrinting.XLActiveSheet(pActiveSheet);
                 
                //발급번호
                mPrinting.XLSetCell(11, 3, pCert.CurrentRow["PRINT_NUM"]);

                //증명서
                //Code(01) : 재직증명서, Code(02) : 경력증명서, Code(03) : 퇴직증명서 
                mPrinting.XLSetCell(2, 2, pCert.CurrentRow["CERTIFICATE_TITLE"]);

                //한글
                mPrinting.XLSetCell(14, 9, pCert.CurrentRow["NAME"]);

                ////한자
                //mPrinting.XLSetCell(15, 13, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));

                //주민번호
                string vREPRE_NUM = "";
                if (pREPRE_FLAG.Equals("Y"))
                    vREPRE_NUM = string.Format("{0}", Convert.ToString(pCert.CurrentRow["REPRE_NUM"]));
                else
                {
                    vREPRE_NUM = string.Format("{0}", Convert.ToString(pCert.CurrentRow["REPRE_NUM"]));
                    vREPRE_NUM = vREPRE_NUM.Substring(0, 6) + "-*******";
                }
                mPrinting.XLSetCell(14, 27, vREPRE_NUM);


                //주소
                mPrinting.XLSetCell(17, 9, pCert.CurrentRow["PERSON_ADDRESS"]);

                //직위
                mPrinting.XLSetCell(20, 9, pCert.CurrentRow["POST_NAME"]);

                //부서
                mPrinting.XLSetCell(20, 27, pCert.CurrentRow["DEPT_NAME"]);
                 
                //재직기간(최종일자) 
                mPrinting.XLSetCell(23, 9, pCert.CurrentRow["RETIRE_DATE"]);
                 
                //담당업무
                mPrinting.XLSetCell(26, 9, pCert.CurrentRow["TASK_DESC"]);

                //용도
                mPrinting.XLSetCell(29, 9, pCert.CurrentRow["REMARK"]);

                ////제출처
                //mPrinting.XLSetCell(39, 9, pCert.CurrentRow["SEND_ORG"]);

                //증명서 설명
                mPrinting.XLSetCell(33, 3, pCert.CurrentRow["CERTIFICATE_REMARK"]);
                  
                //인쇄일자
                mPrinting.XLSetCell(41, 6, pCert.CurrentRow["PRINT_DATE"]);

                //회사명
                mPrinting.XLSetCell(43, 6, pCert.CurrentRow["CORP_NAME"]);

                //회사주소
                mPrinting.XLSetCell(46, 2, pCert.CurrentRow["CORP_ADDRESS"]);

                //대표자명
                mPrinting.XLSetCell(49, 6, pCert.CurrentRow["PRESIDENT_NAME"]);

                //회사 영문(상단)
                mPrinting.XLSetCell(1, 2, pCert.CurrentRow["CORP_NAME_ENG"]);

                if(pHISTORY_FLAG.Equals("Y"))
                {
                    int vLine = m_History_Row;
                    foreach(DataRow vROW in pHistory.CurrentRows)
                    {
                        mPrinting.XLSetCell(vLine, 3, pHistory.CurrentRow["CHARGE_DATE"]);
                        mPrinting.XLSetCell(vLine, 8, pHistory.CurrentRow["END_DATE"]);
                        mPrinting.XLSetCell(vLine, 13, pHistory.CurrentRow["DEPT_NAME"]);
                        mPrinting.XLSetCell(vLine, 23, pHistory.CurrentRow["POST_NAME"]);
                        mPrinting.XLSetCell(vLine, 29, pHistory.CurrentRow["OCPT_NAME"]);
                        vLine++;
                    }
                } 
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }


        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow)
        {
            //object vObject = null;

            try
            {
                mPrinting.XLActiveSheet("Sheet1");
                
                int vIndexDataColumn1 = pGrid.GetColumnToIndex("PRINT_NUM");              //발급번호
                int vIndexDataColumn2 = pGrid.GetColumnToIndex("CERTIFICATE_TITLE");    //증명서
                int vIndexDataColumn3 = pGrid.GetColumnToIndex("NAME");                   //한글
                int vIndexDataColumn4 = pGrid.GetColumnToIndex("CHINESE_NAME");           //한자
                int vIndexDataColumn5 = pGrid.GetColumnToIndex("REPRE_NUM");              //주민등록번호
                int vIndexDataColumn6 = pGrid.GetColumnToIndex("PERSON_ADDRESS");         //주소
                int vIndexDataColumn7 = pGrid.GetColumnToIndex("DEPT_NAME");              //부서
                int vIndexDataColumn8 = pGrid.GetColumnToIndex("POST_NAME");              //직위
                int vIndexDataColumn9 = pGrid.GetColumnToIndex("ORI_JOIN_DATE");          //재직기간(최초일자)
                int vIndexDataColumn10 = pGrid.GetColumnToIndex("RETIRE_DATE");           //재직기간(최종일자)
                int vIndexDataColumn11 = pGrid.GetColumnToIndex("DESCRIPTION");           //용도
                int vIndexDataColumn12 = pGrid.GetColumnToIndex("SEND_ORG");              //제출처
                int vIndexDataColumn13 = pGrid.GetColumnToIndex("CERTIFICATE_REMARK");    //특이사항
                int vIndexDataColumn14 = pGrid.GetColumnToIndex("PRINT_COUNT");           //수량
                int vIndexDataColumn15 = pGrid.GetColumnToIndex("PRINT_DATE");            //인쇄일자
                int vIndexDataColumn16 = pGrid.GetColumnToIndex("CORP_NAME");             //회사명
                int vIndexDataColumn17 = pGrid.GetColumnToIndex("CORP_ADDRESS");          //회사주소
                int vIndexDataColumn18 = pGrid.GetColumnToIndex("PRESIDENT_NAME");        //대표이사
                int vIndexDataColumn19 = pGrid.GetColumnToIndex("WORKING_NAME");          //담당업무

                //발급번호
                mPrinting.XLSetCell(11, 3, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                
                //증명서
                //Code(01) : 재직증명서, Code(02) : 경력증명서, Code(03) : 퇴직증명서
                object vCertificate_Code = pGrid.GetCellValue(pIndexRow, vIndexDataColumn2);
                mPrinting.XLSetCell(2, 2, vCertificate_Code); 
                
                //한글
                mPrinting.XLSetCell(14, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                ////한자
                //mPrinting.XLSetCell(15, 13, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));

                //주민번호
                mPrinting.XLSetCell(14, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));

                //주소
                mPrinting.XLSetCell(17, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));

                //직위
                mPrinting.XLSetCell(20, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn8));

                //부서
                mPrinting.XLSetCell(20, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                  
                //재직기간(최초일자)
                //if(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9) != null)
                //{
                //    object test1 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    //mPrinting.XLSetCell(23, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                //}
                //else
                //    mPrinting.XLSetCell(30, 13, "");

                //재직기간(최종일자)
                //if (pGrid.GetCellValue(pIndexRow, vIndexDataColumn10) != null)
                //{
                //    object test2 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                mPrinting.XLSetCell(23, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                
                //}
                //else
                //    mPrinting.XLSetCell(33, 13, "");

                //담당업무
                mPrinting.XLSetCell(26, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn19));

                //용도
                mPrinting.XLSetCell(29, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn11));

                ////제출처
                //mPrinting.XLSetCell(39, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));

                //적요
                mPrinting.XLSetCell(33, 3, pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));

                ////매 수
                //mPrinting.XLSetCell(45, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn14));

                //인쇄일자
                mPrinting.XLSetCell(41, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn15));                

                //회사명
                mPrinting.XLSetCell(43, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn16));

                //회사주소
                mPrinting.XLSetCell(46, 2, pGrid.GetCellValue(pIndexRow, vIndexDataColumn17));

                //대표자명
                mPrinting.XLSetCell(49, 6, pGrid.GetCellValue(pIndexRow, vIndexDataColumn18));

            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        private void XLContentWrite2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pLINE)
        {
            //object vObject = null;
            int vLINE = pLINE;
            try
            {
                mPrinting.XLActiveSheet("Sheet1");

                int vIndexDataColumn1 = pGrid.GetColumnToIndex("CHARGE_DATE");        //재직시작일
                int vIndexDataColumn5 = pGrid.GetColumnToIndex("END_DATE");           //재직종료일
                int vIndexDataColumn2 = pGrid.GetColumnToIndex("DEPT_NAME");          //소속 (부서)
                int vIndexDataColumn3 = pGrid.GetColumnToIndex("POST_NAME");          //직위
                int vIndexDataColumn4 = pGrid.GetColumnToIndex("OCPT_NAME");          //담당업무 (직무) 


                //재직시작일
                mPrinting.XLSetCell(pLINE, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                //재직종료일
                mPrinting.XLSetCell(pLINE, 10, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));

                //소속 --부서 
                mPrinting.XLSetCell(pLINE, 15, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));

                //직위
                mPrinting.XLSetCell(pLINE, 25, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                //직무
                mPrinting.XLSetCell(pLINE, 31, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));


                pLINE = pLINE + 1;
                //재직기간(최초일자)
                //if(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9) != null)
                //{
                //    object test1 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                //mPrinting.XLSetCell(23, 9, pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                //}
                //else
                //    mPrinting.XLSetCell(30, 13, "");

                //재직기간(최종일자)
                //if (pGrid.GetCellValue(pIndexRow, vIndexDataColumn10) != null)
                //{
                //    object test2 = ConvertDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));


            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }


        #endregion;

        #region ----- Excel Wirte Methods ----

        public int XLWirte(InfoSummit.Win.ControlAdv.ISDataAdapter pCert, InfoSummit.Win.ControlAdv.ISDataAdapter pHistory
                        , int nPrintTotalCnt, string pPeriodFrom
                        , string pUserName, string pPRINT_TYPE, string pREPRE_FLAG, string pHISTORY_FLAG
                        , string pSTAMP_FLAG, string pClientFile, float pSIZE_W, float pSIZE_H, float pLOC_X, float pLOC_Y)
        {
            string vMessageText = string.Empty;
            string vSheet_Source = string.Empty;
            m_Current_Row = 1;         
            try
            {
                if(pHISTORY_FLAG.Equals("Y"))
                {
                    vSheet_Source = m_SheetSource2; 
                }
                else
                {
                    vSheet_Source = m_SheetSource1; 
                }
                XLContentWrite(vSheet_Source, pCert, pREPRE_FLAG, pHISTORY_FLAG, pHistory);
                for (int nPrintCnt = 0; nPrintCnt < nPrintTotalCnt; nPrintCnt++)
                {
                    m_Current_Row = CopyAndPaste(mPrinting, m_Current_Row, vSheet_Source
                                                , pSTAMP_FLAG, pClientFile, pSIZE_W, pSIZE_H, pLOC_X, pLOC_Y); 
                }
            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            //sheet 삭제//
            mPrinting.XLDeleteSheet(m_SheetSource1);
            mPrinting.XLDeleteSheet(m_SheetSource2); 
            return nPrintTotalCnt;
        }
         
        #endregion;


        #region ----- Excel Copy&Paste Methods ----


        //첫번째 페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCurrentRow, string pSourceSheet)
        { 
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceSheet);
            object vRangeSource = pPrinting.XLGetRange(m_Copy_StartRow, m_Copy_StartCol, m_Copy_EndRow, m_Copy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(m_SheetPrint);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, m_Copy_StartCol, (m_Copy_EndRow * (m_PageNumber + 1)) + 1, m_Copy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            int vCopy_EndRow = pCurrentRow + m_Copy_EndRow;
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopy_EndRow));

            m_PageNumber++; //페이지 번호

            return pCurrentRow + m_Copy_EndRow;
        }

        //첫번째 페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCurrentRow, string pSourceSheet, string pSeal_Stamp)
        {
            if (pSeal_Stamp == "N")
            {
                mPrinting.XLDeleteBarCode(pIndexImage: 1);
            }

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceSheet);
            object vRangeSource = pPrinting.XLGetRange(m_Copy_StartRow, m_Copy_StartCol, m_Copy_EndRow, m_Copy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(m_SheetPrint);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, m_Copy_StartCol, pCurrentRow + m_Copy_EndRow, m_Copy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            int vCopy_EndRow = pCurrentRow + m_Copy_EndRow;
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopy_EndRow));

            m_PageNumber++; //페이지 번호

            return pCurrentRow + m_Copy_EndRow;
        }

        //첫번째 페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCurrentRow, string pSourceSheet
                                , string pSeal_Stamp, string pImageFile
                                , float pSize_W, float pSize_H, float pLoc_X, float pLoc_Y)
        {
            pPrinting.XLActiveSheet(pSourceSheet);
            if (pSeal_Stamp == "Y")
            {
                //----------------------------------------[ 증명사진 출력 부분 ]------------------------------------------
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
                    //
                }

                try
                {
                    System.Drawing.SizeF vSize = new System.Drawing.SizeF(pSize_W, pSize_H);
                    System.Drawing.PointF vPoint = new System.Drawing.PointF(pLoc_X, pLoc_Y);
                    mPrinting.XLBarCode(pImageFile, vSize, vPoint);
                    //--------------------------------------------------------------------------------------------------------
                }
                catch
                {
                    //
                }

            }
            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            object vRangeSource = pPrinting.XLGetRange(m_Copy_StartRow, m_Copy_StartCol, m_Copy_EndRow, m_Copy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(m_SheetPrint);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, m_Copy_StartCol, pCurrentRow + m_Copy_EndRow, m_Copy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            int vCopy_EndRow = pCurrentRow + m_Copy_EndRow;
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopy_EndRow));

            m_PageNumber++; //페이지 번호

            return pCurrentRow + m_Copy_EndRow;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPrinting(pPageSTART, pPageEND, 1);
                //mPrinting.XLPrinting(pPageSTART, pPageEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        public void PreView(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
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