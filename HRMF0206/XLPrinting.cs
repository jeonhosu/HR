using System;

namespace HRMF0206
{
    /// <summary>
    /// XLPrint Class를 이용해 Report물 제어 
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

        private int mPositionPrintLineSTART = 1; //내용 출력시 엑셀 시작 행 위치 지정
        private int[] mIndexXLWriteColumn = new int[0] { }; //엑셀에 출력할 열 위치 지정

        private int mMaxIncrement = 41; //실제 출력되는 행의 시작부터 끝 행의 범위
        private int mSumPrintingLineCopy = 1; //엑셀의 선택된 쉬트에 복사되어질 시작 행 위치 및 누적 행 값
        private int mMaxIncrementCopy = 67; //반복 복사되어질 행의 최대 범위

        private int mXLColumnAreaSTART = 1; //복사되어질 쉬트의 폭, 시작열
        private int mXLColumnAreaEND = 45;  //복사되어질 쉬트의 폭, 종료열

        //복사 영역//
        private int mSTART_COL = 1;
        private int mSTART_ROW = 1;
        private int mEND_COL = 45;
        private int mEND_ROW = 41;

        private int mCURRENT_ROW = 0;

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

        #region ----- Define Print Column Methods ----

        private void XLDefinePrintColumn(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            try
            {
                //Grid의 [Edit] 상의 [DataColumn] 열에 있는 열 이름을 지정 하면 된다.
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

                //엑셀에 출력될 열 위치 지정
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

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        //string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                        string vTextDateTimeLong = vDateTime.ToString("yyyy년 MM월 dd일", null);
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
            // 제목 및 기본사항 항목명 출력 부분
            //======================================================================================
            //제목
            mPrinting.XLSetCell(5, 13, "인   사   기   록   표");
            //기본사항
            mPrinting.XLSetCell(10, 3, "기   본   사   항");
            //성명
            mPrinting.XLSetCell(12, 11, "성      명");
            //직군
            mPrinting.XLSetCell(12, 22, "직      군");
            //급여구분
            mPrinting.XLSetCell(12, 33, "급여구분");
            //부서
            mPrinting.XLSetCell(14, 11, "부      서");
            //직책
            mPrinting.XLSetCell(14, 22, "직      책");
            //주민등록번호
            mPrinting.XLSetCell(14, 33, "주민등록번호");
            //사번
            mPrinting.XLSetCell(16, 11, "사      번");
            //직위
            mPrinting.XLSetCell(16, 22, "직      위");
            //퇴사일자
            mPrinting.XLSetCell(16, 33, "퇴사일자");
            //입사일자
            mPrinting.XLSetCell(18, 11, "입사일자");
            //직무
            mPrinting.XLSetCell(18, 22, "직      무");
            //계좌
            mPrinting.XLSetCell(18, 33, "계      좌");
            //전화번호
            mPrinting.XLSetCell(20, 11, "전화번호");
            //이메일
            mPrinting.XLSetCell(20, 22, "이 메 일");
            //노조가입
            mPrinting.XLSetCell(20, 33, "노조가입");
            //======================================================================================
            // 책정임금 항목명
            //======================================================================================
            //책정임금
            mPrinting.XLSetCell(44, 25, "책정임금");
            //적용기간
            mPrinting.XLSetCell(44, 27, "적용기간");
            //======================================================================================
            // 학력사항 항목명
            //======================================================================================
            //학력사항
            mPrinting.XLSetCell(23, 3, "학력사항");
            //년월
            mPrinting.XLSetCell(23, 5, "년 월");
            //출신교
            mPrinting.XLSetCell(23, 10, "출 신 교");
            //학력
            mPrinting.XLSetCell(23, 16, "학 력");
            //전공
            mPrinting.XLSetCell(23, 19, "전 공");
            //======================================================================================
            // 가족사항 항목명
            //======================================================================================
            //가족사항
            mPrinting.XLSetCell(23, 25, "가족사항");
            //관계
            mPrinting.XLSetCell(23, 27, "관 계");
            //성명
            mPrinting.XLSetCell(23, 30, "성 명");
            //생년월일
            mPrinting.XLSetCell(23, 34, "생년월일");
            //학력
            mPrinting.XLSetCell(23, 38, "학 력");
            //근무처
            mPrinting.XLSetCell(23, 41, "근 무 처");
            //======================================================================================
            // 자격사항 항목명
            //======================================================================================
            //자격/면허
            mPrinting.XLSetCell(32, 3, "자격/면허");
            //자격증명
            mPrinting.XLSetCell(32, 5, "자격증명");
            //등급
            mPrinting.XLSetCell(32, 12, "등급");
            //취득일
            mPrinting.XLSetCell(32, 17, "취득일");
            //======================================================================================
            // 경력사항 항목명
            //======================================================================================
            //경력사항
            mPrinting.XLSetCell(32, 25, "경력사항");
            //근무기간
            mPrinting.XLSetCell(32, 27, "근무기간");
            //근무처
            mPrinting.XLSetCell(32, 33, "근무처");
            //직급
            mPrinting.XLSetCell(32, 38, "직급");
            //담당업무
            mPrinting.XLSetCell(32, 41, "담당업무");
            //======================================================================================
            // 어학사항 항목명
            //======================================================================================
            //어학
            mPrinting.XLSetCell(38, 3, "어 학");
            //어학구분
            mPrinting.XLSetCell(38, 5, "어학구분");
            //종류
            mPrinting.XLSetCell(38, 11, "종 류");
            //등급
            mPrinting.XLSetCell(38, 17, "등급");
            //점수
            mPrinting.XLSetCell(38, 20, "점 수");
            //======================================================================================
            // 표창/징계사항 항목명
            //======================================================================================
            //표창/징계
            mPrinting.XLSetCell(38, 25, "표창/징계");
            //상벌일자
            mPrinting.XLSetCell(38, 27, "상벌일자");
            //상벌구분
            mPrinting.XLSetCell(38, 33, "상벌구분");
            //종류
            mPrinting.XLSetCell(38, 37, "종류");
            //내용
            mPrinting.XLSetCell(38, 41, "내용");
            //======================================================================================
            // 교육사항 항목명
            //======================================================================================
            //교육
            mPrinting.XLSetCell(44, 3, "교 육");
            //교육구분
            mPrinting.XLSetCell(44, 5, "교육구분");
            //기간
            mPrinting.XLSetCell(44, 11, "기 간");
            //교육명
            mPrinting.XLSetCell(44, 18, "교육명");
            //======================================================================================
            // 발령사항 항목명
            //======================================================================================
            //발령이력
            mPrinting.XLSetCell(50, 3, "발 령 이 력");
            //발령일자
            mPrinting.XLSetCell(50, 5, "발령일자");
            //발령
            mPrinting.XLSetCell(50, 12, "발 령");
            //부서
            mPrinting.XLSetCell(50, 18, "부 서");
            //직책
            mPrinting.XLSetCell(50, 23, "직 책");
            //직급
            mPrinting.XLSetCell(50, 27, "직 급");
            //직위
            mPrinting.XLSetCell(50, 31, "직 위");
            //호봉
            mPrinting.XLSetCell(50, 35, "호 봉");
            //비고
            mPrinting.XLSetCell(50, 39, "비 고");
            //======================================================================================
            // 신체/병력사항 항목명
            //======================================================================================
            //병력
            mPrinting.XLSetCell(28, 3, "병력");
            //장애
            mPrinting.XLSetCell(28, 16, "장애");
            //신체
            mPrinting.XLSetCell(29, 3, "신체");
            //병역
            mPrinting.XLSetCell(29, 16, "병역");
            //주소
            mPrinting.XLSetCell(30, 3, "주소");
            //======================================================================================
            // 용지 하단의 출력 정보 항목명
            //======================================================================================
            //출력자
            mPrinting.XLSetCell(65, 27, "출력자 : ");
            //출력일자
            mPrinting.XLSetCell(65, 37, "출력일자 : ");
        }

        public void ReportTitle2()
        {
            //======================================================================================
            // 제목 및 기본사항 항목명 출력 부분
            //======================================================================================
            //제목
            mPrinting.XLSetCell(1, 12, "[인 사 기 록 카 드]");

            //기본사항
            mPrinting.XLSetCell(4, 2, "기 본 사 항");
            mPrinting.XLSetCell(5, 10, "부 서 명");
            mPrinting.XLSetCell(5, 28, "고용형태");
            mPrinting.XLSetCell(6, 10, "직    급");
            mPrinting.XLSetCell(6, 28, "직    책");
            mPrinting.XLSetCell(7, 10, "성    명");
            mPrinting.XLSetCell(7, 28, "급여형태");
            mPrinting.XLSetCell(8, 10, "생년월일");
            mPrinting.XLSetCell(8, 28, "나이(만)");
            mPrinting.XLSetCell(9, 10, "입 사 일");
            mPrinting.XLSetCell(9, 28, "퇴 사 일");
            mPrinting.XLSetCell(10, 10, "입사유형");
            mPrinting.XLSetCell(10, 28, "연 락 처");
            mPrinting.XLSetCell(11, 28, "현 주 소");

            //======================================================================================
            // 학력사항
            //======================================================================================
            mPrinting.XLSetCell(12, 2, "학력사항");

            mPrinting.XLSetCell(13, 2, "년월");
            mPrinting.XLSetCell(13, 9, "학교명");
            mPrinting.XLSetCell(13, 15, "졸업구분");
            mPrinting.XLSetCell(13, 18, "소재지");

            //======================================================================================
            //경력사항
            //======================================================================================
            mPrinting.XLSetCell(12, 23, "경력사항");

            mPrinting.XLSetCell(13, 23, "년월");
            mPrinting.XLSetCell(13, 30, "회사명");
            mPrinting.XLSetCell(13, 37, "직급");
            mPrinting.XLSetCell(13, 40, "담당업무");

            //======================================================================================
            // 인사평가
            //======================================================================================
            mPrinting.XLSetCell(19, 2, "인사평가");

            mPrinting.XLSetCell(20, 2, "평가년도");
            mPrinting.XLSetCell(20, 7, "평가등급");
            mPrinting.XLSetCell(20, 14, "비고"); 

            //======================================================================================
            // 발령사항
            //======================================================================================
            mPrinting.XLSetCell(19, 23, "발령사항");

            mPrinting.XLSetCell(20, 23, "발령일자");
            mPrinting.XLSetCell(20, 27, "발령사유");
            mPrinting.XLSetCell(20, 34, "부서명");
            mPrinting.XLSetCell(20, 40, "직급");

            //======================================================================================
            // 자격/어학
            //======================================================================================            
            mPrinting.XLSetCell(25, 2, "자격/어학");

            mPrinting.XLSetCell(26, 2, "명칭");
            mPrinting.XLSetCell(26, 8, "등급");
            mPrinting.XLSetCell(26, 13, "점수");
            mPrinting.XLSetCell(26, 18, "취득일자");

            //======================================================================================
            // 교육사항
            //======================================================================================
            mPrinting.XLSetCell(31, 2, "교육사항");
            
            mPrinting.XLSetCell(32, 2, "교육명");
            mPrinting.XLSetCell(38, 9, "교육기간");
            mPrinting.XLSetCell(38, 16, "시행처");

            //======================================================================================
            // 기타사항
            //======================================================================================
            mPrinting.XLSetCell(36, 2, "기타사항");
            
            mPrinting.XLSetCell(37, 2, "구분");
            mPrinting.XLSetCell(37, 6, "내용");
            mPrinting.XLSetCell(37, 12, "비고");

            mPrinting.XLSetCell(38, 2, "병역");
            mPrinting.XLSetCell(39, 2, "장애");
            mPrinting.XLSetCell(40, 2, "보훈여부");

            //======================================================================================
            // 상벌사항
            //======================================================================================
            mPrinting.XLSetCell(36, 23, "상벌사항");

            //상벌일자
            mPrinting.XLSetCell(37, 23, "상벌일자");
            //상벌구분
            mPrinting.XLSetCell(37, 27, "상벌구분");
            //종류
            mPrinting.XLSetCell(37, 30, "종류");
            //내용
            mPrinting.XLSetCell(37, 37, "내용");
        }

        #endregion;

        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow, int pTotalRow, int pCnt, string pPrintDateTime, string pUserName)
        {
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                if (pCnt == 1)
                {   
                    // 기본 정보1
                    int vIndexDataColumn1  = pGrid.GetColumnToIndex("NAME");            // 성명
                    int vIndexDataColumn2  = pGrid.GetColumnToIndex("JOB_CLASS_NAME");  // 직군
                    int vIndexDataColumn3  = pGrid.GetColumnToIndex("DEPT_NAME");       // 부서    
                    int vIndexDataColumn4  = pGrid.GetColumnToIndex("ABIL_NAME");       // 직책    
                    int vIndexDataColumn5  = pGrid.GetColumnToIndex("REPRE_NUM");       // 주민번호
                    int vIndexDataColumn6  = pGrid.GetColumnToIndex("PERSON_NUM");      // 사번    
                    int vIndexDataColumn7  = pGrid.GetColumnToIndex("D_POST_NAME");       // 직위    
                    int vIndexDataColumn8  = pGrid.GetColumnToIndex("RETIRE_DATE");     // 퇴사일자
                    int vIndexDataColumn9  = pGrid.GetColumnToIndex("JOIN_DATE");       // 입사일자
                    int vIndexDataColumn10 = pGrid.GetColumnToIndex("OCPT_NAME");       // 직무
                    int vIndexDataColumn11 = pGrid.GetColumnToIndex("PRSN_ADDR1");      // 주소1
                    int vIndexDataColumn21 = pGrid.GetColumnToIndex("PRSN_ADDR2");      // 주소2
                    int vIndexDataColumn12 = pGrid.GetColumnToIndex("EMAIL");           // 이메일
                    int vIndexDataColumn13 = pGrid.GetColumnToIndex("TELEPHON_NO");     // 전화번호
                    int vIndexDataColumn14 = pGrid.GetColumnToIndex("LABOR_UNION_YN");  // 노조가입

                    //성명
                    mPrinting.XLSetCell(12, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //직군
                    mPrinting.XLSetCell(12, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //부서
                    mPrinting.XLSetCell(14, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //직책
                    mPrinting.XLSetCell(14, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //주민번호
                    mPrinting.XLSetCell(14, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    //사번
                    mPrinting.XLSetCell(16, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    //직위
                    mPrinting.XLSetCell(16, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                    //퇴사일자
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
                    //입사일자
                    DateTime dJoinDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn9));
                    object vJoinDate = dJoinDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(18, 16, vJoinDate);
                    //직무
                    mPrinting.XLSetCell(18, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn10));
                    //주소
                    object vAddress = string.Format("{0} {1}", pGrid.GetCellValue(pIndexRow, vIndexDataColumn11), pGrid.GetCellValue(pIndexRow, vIndexDataColumn21));
                    mPrinting.XLSetCell(30, 5, vAddress);
                    //이메일
                    mPrinting.XLSetCell(20, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn12));
                    //전화번호
                    mPrinting.XLSetCell(20, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn13));
                    //노조가입
                    object vLaborUnion = pGrid.GetCellValue(pIndexRow, vIndexDataColumn14);
                    if (vLaborUnion.ToString() == "N")
                    {
                        mPrinting.XLSetCell(20, 38, "미가입");
                    }
                    else if (vLaborUnion.ToString() == "Y")
                    {
                        mPrinting.XLSetCell(20, 38, "가입");
                    }
                    else
                    {
                        mPrinting.XLSetCell(20, 38, ""); 
                    }
                }
                else if (pCnt == 2)
                {                 
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("PAYMENT_DATE");   // 적용기간
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("BANK_NAME");      // 은행명  
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BANK_ACCOUNTS");  // 계좌번호
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("PAY_TYPE_NAME");  // 급여구분

                    //적용기간
                    mPrinting.XLSetCell(45 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //은행명
                    mPrinting.XLSetCell(18, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //계좌번호
                    mPrinting.XLSetCell(19, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //급여구분
                    mPrinting.XLSetCell(12, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 3)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("SCHOLARSHIP_TYPE_NAME"); // 학력         
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("GRADUATION_YYYYMM");     // 졸업일자
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("SCHOOL_NAME");           // 출신교
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("SPECIAL_STUDY_NAME");    // 전공                

                    //학력
                    mPrinting.XLSetCell(24 + pIndexRow, 16, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //졸업일자
                    mPrinting.XLSetCell(24 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //출신교
                    mPrinting.XLSetCell(24 + pIndexRow, 10, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //전공 
                    mPrinting.XLSetCell(24 + pIndexRow, 19, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 4)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("FAMILY_NAME");    // 성명    
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("RELATION_NAME");  // 관계    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BIRTHDAY");       // 생년월일
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("COMPANY_NAME");   // 회사명 
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("END_SCH_NAME");   // 학력

                    //성명 
                    mPrinting.XLSetCell(24 + pIndexRow, 30, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //관계
                    mPrinting.XLSetCell(24 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //생년월일 
                    DateTime vBirthday = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    string sBirthday = vBirthday.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(24 + pIndexRow, 34, sBirthday);
                    //회사명
                    mPrinting.XLSetCell(24 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //학력
                    mPrinting.XLSetCell(24 + pIndexRow, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                }
                else if (pCnt == 5)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("LICENSE_NAME");         // 자격증명
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("LICENSE_GRADE_NAME");   // 자격등급
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("LICENSE_DATE");         // 취득일자

                    //자격증명
                    mPrinting.XLSetCell(33 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //자격등급
                    mPrinting.XLSetCell(33 + pIndexRow, 12, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //취득일자
                    mPrinting.XLSetCell(33 + pIndexRow, 17, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                }
                else if (pCnt == 6)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("COMPANY_NAME");   // 근무처  
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("POST_NAME");      // 직급    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("JOB_NAME");       // 담당업무
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("START_DATE");     // 입사일  
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("END_DATE");       // 퇴사일  

                    //근무처
                    mPrinting.XLSetCell(33 + pIndexRow, 33, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //직급
                    mPrinting.XLSetCell(33 + pIndexRow, 38, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //담당업무
                    mPrinting.XLSetCell(33 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                    //입사일                   
                    DateTime dStartDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow,vIndexDataColumn4));
                    string sStartDate = dStartDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    //퇴사일
                    DateTime dEndDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    string sEndDate = dEndDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);

                    object vStartEndDate = sStartDate + " ~ " + sEndDate;
                    mPrinting.XLSetCell(33 + pIndexRow, 27, vStartEndDate);
                }
                else if (pCnt == 7)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("LANGUAGE_NAME");  // 어학구분
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("EXAM_NAME");      // 어학종류
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("EXAM_LEVEL");     // 등급    
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("SCORE");          // 점수    

                    //어학구분
                    mPrinting.XLSetCell(39 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //어학종류
                    mPrinting.XLSetCell(39 + pIndexRow, 11, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //등급
                    mPrinting.XLSetCell(39 + pIndexRow, 17, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //점수
                    mPrinting.XLSetCell(39 + pIndexRow, 20, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 8)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("RP_TYPE_NAME");    // 상벌구분
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("RP_NAME");         // 상벌사항
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("RP_DATE");         // 상벌일자
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("RP_DESCRIPTION");  // 상벌내용

                    //상벌구분
                    mPrinting.XLSetCell(39 + pIndexRow, 33, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    //상벌종류
                    mPrinting.XLSetCell(39 + pIndexRow, 37, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //상벌일자
                    DateTime dRP_Date = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    string sRP_Date = dRP_Date.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(39 + pIndexRow, 27, sRP_Date);

                    //상벌내용
                    mPrinting.XLSetCell(39 + pIndexRow, 41, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 9)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("START_DATE");      // 시작일자
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("END_DATE");        // 종료일자
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("EDU_ORG");         // 교육구분
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("EDU_CURRICULUM");  // 교육과목

                    //시작일자
                    DateTime dStartDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    string sStartDate = dStartDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    //종료일자
                    DateTime dEndDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    string sEndDate = dEndDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);

                    object vStartEndDate = sStartDate + " ~ " + sEndDate;
                                        
                    //시작일자~종료일자
                    mPrinting.XLSetCell(45 + pIndexRow, 11, vStartEndDate);
                    //교육구분
                    mPrinting.XLSetCell(45 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //교육과목
                    mPrinting.XLSetCell(45 + pIndexRow, 18, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 10)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("CHARGE_DATE");    // 발령일자
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("CHARGE_NAME");    // 발령    
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("DESCRIPTION");    // 비고    
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("DEPT_NAME");      // 부서    
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("POST_NAME");      // 직위    
                    int vIndexDataColumn6 = pGrid.GetColumnToIndex("ABIL_NAME");      // 직책    
                    int vIndexDataColumn7 = pGrid.GetColumnToIndex("PAY_GRADE_NAME"); // 직급    

                    //발령일자
                    DateTime dChargeDate = Convert.ToDateTime(pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));
                    object vChargeDate = dChargeDate.ToString("yyyy-MM-dd", null).Replace("0001-01-01", null);
                    mPrinting.XLSetCell(51 + pIndexRow, 5, vChargeDate);
                    //발령
                    mPrinting.XLSetCell(51 + pIndexRow, 12, pGrid.GetCellValue(pIndexRow, vIndexDataColumn2));
                    //비고
                    mPrinting.XLSetCell(51 + pIndexRow, 39, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));
                    //부서
                    mPrinting.XLSetCell(51 + pIndexRow, 18, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //직위
                    mPrinting.XLSetCell(51 + pIndexRow, 31, pGrid.GetCellValue(pIndexRow, vIndexDataColumn5));
                    //직책
                    mPrinting.XLSetCell(51 + pIndexRow, 23, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));
                    //직급
                    mPrinting.XLSetCell(51 + pIndexRow, 27, pGrid.GetCellValue(pIndexRow, vIndexDataColumn7));
                }
                else if (pCnt == 11)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("ARMY_KIND_NAME");     // 군별
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("ARMY_GRADE_NAME");    // 계급 
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("ARMY_END_TYPE_NAME"); // 전역구분
                    //int vIndexDataColumn4 = pGrid.GetColumnToIndex("DESCRIPTION");      // 병력

                    // 병역항목 - 군별, 계급, 전역구분
                    object vArmyInfo = pGrid.GetCellValue(pIndexRow, vIndexDataColumn1).ToString() + ", "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn2).ToString() + ", "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn3).ToString();

                    mPrinting.XLSetCell(29 + pIndexRow, 19, vArmyInfo);

                    //병력
                    //mPrinting.XLSetCell(28 + pIndexRow, 5, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                }
                else if (pCnt == 12)
                {
                    int vIndexDataColumn1 = pGrid.GetColumnToIndex("HEIGHT");         // 키    
                    int vIndexDataColumn2 = pGrid.GetColumnToIndex("WEIGHT");         // 몸무게
                    int vIndexDataColumn3 = pGrid.GetColumnToIndex("BLOOD_NAME");     // 혈액형
                    int vIndexDataColumn4 = pGrid.GetColumnToIndex("DISABLED_NAME");  // 장애
                    int vIndexDataColumn5 = pGrid.GetColumnToIndex("DESCRIPTION");    // 병력
                    
                    // 신체항목 - 키, 몸무게, 혈액형
                    object vBodyInfo = pGrid.GetCellValue(pIndexRow, vIndexDataColumn1).ToString() + "cm, "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn2).ToString() + "kg, "
                                     + pGrid.GetCellValue(pIndexRow, vIndexDataColumn3).ToString();

                    mPrinting.XLSetCell(29 + pIndexRow, 5, vBodyInfo);

                    //장애
                    mPrinting.XLSetCell(28 + pIndexRow, 19, pGrid.GetCellValue(pIndexRow, vIndexDataColumn4));
                    //병력
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
            int vTotalRow = pGrid.RowCount; // Grid의 총 행수

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

                if (pCnt == 12) // 12번째 마지막 Grid일 경우,
                {
                    //----------------------------------------[ 증명사진 출력 부분 ]------------------------------------------
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

                    //인사내역 문서에 항목명을 출력해주는 함수 호출
                    ReportTitle();

                    //문서 하단에 출력 정보 표시
                    mPrinting.XLSetCell(65, 31, pUserName);
                    mPrinting.XLSetCell(65, 41, pPrintDateTime);

                    //[Sheet2]내용을 [Sheet1]에 붙여넣기
                    mSumPrintingLineCopy = CopyAndPaste(mSumPrintingLineCopy);

                    //-------------------------------------------------------------------------------------------------------
                    // 페이지 내용 삭제 부분
                    // (SourceTab1에 데이터 출력 -> Destination에 복사 -> SourceTab1 데이터 삭제 후, 다음 데이터 출력 
                    //-------------------------------------------------------------------------------------------------------
                    mPrinting.XLActiveSheet("SourceTab1");
                    int vStartRow = mPositionPrintLineSTART; //시작 행 위치 부터
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
                int vIDX_Col0 = pGrid.GetColumnToIndex("REMARK");               // 적요
                int vIDX_Col0_1 = pGrid.GetColumnToIndex("PERSON_NUM");         // 사원번호

                // 기본 정보1
                int vIDX_Col1 = pGrid.GetColumnToIndex("DEPT_NAME");            // 부서
                int vIDX_Col2 = pGrid.GetColumnToIndex("CONTRACT_TYPE_NAME");   // 고용형태
                int vIDX_Col3 = pGrid.GetColumnToIndex("D_POST_NAME");            // 직위 
                int vIDX_Col4 = pGrid.GetColumnToIndex("ABIL_NAME");            // 직책    
                int vIDX_Col5 = pGrid.GetColumnToIndex("NAME");                 // 성명
                int vIDX_Col6 = pGrid.GetColumnToIndex("PAY_TYPE_NAME");        // 급여형태
                int vIDX_Col7 = pGrid.GetColumnToIndex("BIRTHDAY");             // 생년월일
                int vIDX_Col8 = pGrid.GetColumnToIndex("REAL_AGE");             // 만나이
                int vIDX_Col9 = pGrid.GetColumnToIndex("JOIN_DATE");            // 입사일자
                int vIDX_Col10 = pGrid.GetColumnToIndex("RETIRE_DATE");         // 퇴사일자
                int vIDX_Col11 = pGrid.GetColumnToIndex("JOIN_NAME");           // 입사유형
                int vIDX_Col12 = pGrid.GetColumnToIndex("HP_PHONE_NO");         // 연락처
                int vIDX_Col13 = pGrid.GetColumnToIndex("PRSN_ADDR");           // 현주소

                int vIDX_Col14 = pGrid.GetColumnToIndex("ARMY_END_TYPE_NAME");  // 병역내용
                int vIDX_Col15 = pGrid.GetColumnToIndex("ARMY_PERIOD");         // 병역기간
                int vIDX_Col16 = pGrid.GetColumnToIndex("DISABLED_NAME");       // 장애내용
                int vIDX_Col17 = pGrid.GetColumnToIndex("DISABLED_TYPE");       // 장애등급
                int vIDX_Col18 = pGrid.GetColumnToIndex("BOHUN_NAME");          // 보훈대상
                int vIDX_Col19 = pGrid.GetColumnToIndex("BOHUN_RELATION_NAME"); // 보훈관계

                //적요
                mPrinting.XLSetCell(3, 35, pGrid.GetCellValue(pRow, vIDX_Col0));
                mPrinting.XLSetCell(4, 35, string.Format("[{0}]", pGrid.GetCellValue(pRow, vIDX_Col0_1)));
 
                //부서
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col1));
                //고용형태
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col2));

                //--//
                vRow++;

                //직위
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col3));
                //직책
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col4));

                //--//
                vRow++;

                //성명
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col5));
                //급여형태
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col6));

                //--//
                vRow++;

                //생년월일
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col7));
                //만나이
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col8));

                //--//
                vRow++;

                //입사일자
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col9));
                //퇴사일자
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col10));

                //--//
                vRow++;

                //입사유형
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col11));
                //연락처
                mPrinting.XLSetCell(vRow, 33, pGrid.GetCellValue(pRow, vIDX_Col12));

                //--//
                vRow++;

                //현주소
                mPrinting.XLSetCell(vRow, 15, pGrid.GetCellValue(pRow, vIDX_Col13));

                //기타사항
                vRow = 38;
                //병역 내용
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col14));
                //병역 비고
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col15));

                //--//
                vRow++;

                //장애 내용
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col16));
                //장애 비고
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col17));

                //--//
                vRow++;

                //보훈 내용
                mPrinting.XLSetCell(vRow, 6, pGrid.GetCellValue(pRow, vIDX_Col18));
                //보훈 비고
                mPrinting.XLSetCell(vRow, 12, pGrid.GetCellValue(pRow, vIDX_Col19));
            }
            catch
            {
                return 1;
            }

            //----------------------------------------[ 증명사진 출력 부분 ]------------------------------------------
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

                //문서 하단에 출력 정보 표시
                mPrinting.XLSetCell(41, 1, pUserName);
                mPrinting.XLSetCell(41, 30, pPrintDateTime);
            }
            catch
            {
                return 1;
            }

            //[Sheet2]내용을 [Sheet1]에 붙여넣기
            mSumPrintingLineCopy = CopyAndPaste2(mSumPrintingLineCopy);

            //-------------------------------------------------------------------------------------------------------
            // 페이지 내용 삭제 부분
            // (SourceTab1에 데이터 출력 -> Destination에 복사 -> SourceTab1 데이터 삭제 후, 다음 데이터 출력 
            //-------------------------------------------------------------------------------------------------------
            try
            {
                mPrinting.XLActiveSheet("SourceTab1");
                //원본 초기화//           
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
            //학력사항 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 14;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("SCHOLARSHIP_PERIOD"); // 년월         
                int vIDX_Col2 = pGrid.GetColumnToIndex("SCHOOL_NAME");     // 학교명
                int vIDX_Col3 = pGrid.GetColumnToIndex("GRADUATION_TYPE_NAME");           // 졸업구분
                int vIDX_Col4 = pGrid.GetColumnToIndex("ADDRESS");    // 소재지            
                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //학력
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    //졸업일자
                    mPrinting.XLSetCell(vROW + r, 9, pGrid.GetCellValue(r, vIDX_Col2));
                    //출신교
                    mPrinting.XLSetCell(vROW + r, 15, pGrid.GetCellValue(r, vIDX_Col3));
                    //전공 
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
            //경력사항 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 14;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("CAREER_PERIOD");        // 년월         
                int vIDX_Col2 = pGrid.GetColumnToIndex("COMPANY_NAME");         // 회사명
                int vIDX_Col3 = pGrid.GetColumnToIndex("POST_NAME");            // 직급
                int vIDX_Col4 = pGrid.GetColumnToIndex("JOB_NAME");             // 담당업무

                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //년월
                    mPrinting.XLSetCell(vROW + r, 23, pGrid.GetCellValue(r, vIDX_Col1));
                    //회사명
                    mPrinting.XLSetCell(vROW + r, 30, pGrid.GetCellValue(r, vIDX_Col2));
                    //직급
                    mPrinting.XLSetCell(vROW + r, 37, pGrid.GetCellValue(r, vIDX_Col3));
                    //담당업무 
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
            //인사평가 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 21;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("RESULT_YYYY");          // 평가년도         
                int vIDX_Col2 = pGrid.GetColumnToIndex("RES_LVEL");             // 평가등급
                //int vIDX_Col3 = pGrid.GetColumnToIndex("RES_SCORE");          // 비고
                int vIDX_Col4 = pGrid.GetColumnToIndex("DESCRIPTION");          // 비고
              
                for (int r = 0; r < pGrid.RowCount; r++)
                {
                    //년월
                    mPrinting.XLSetCell(vROW + r, 2, pGrid.GetCellValue(r, vIDX_Col1));
                    //회사명
                    mPrinting.XLSetCell(vROW + r, 7, pGrid.GetCellValue(r, vIDX_Col2));
                    //직급
                    //mPrinting.XLSetCell(vROW + r, 37, pGrid.GetCellValue(r, vIDX_Col3));
                    //담당업무 
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
            //인사발령
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 21;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("CHARGE_DATE");          // 발령일자         
                int vIDX_Col2 = pGrid.GetColumnToIndex("CHARGE_NAME");          // 발령명칭
                int vIDX_Col3 = pGrid.GetColumnToIndex("DEPT_NAME");            // 부서
                int vIDX_Col4 = pGrid.GetColumnToIndex("POST_NAME");            // 직급

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
            //자격/어학 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 27;  //인쇄 위치 
            object vOBJECT;
            string vSTRING;
            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("LICENSE_NAME");         // 명칭         
                int vIDX_Col2 = pGrid.GetColumnToIndex("LICENSE_GRADE_NAME");   // 등급
                int vIDX_Col3 = pGrid.GetColumnToIndex("LICENSE_SCORE");        // 점수
                int vIDX_Col4 = pGrid.GetColumnToIndex("LICENSE_DATE");         // 취득일자

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
            //교육사항
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 33;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("EDU_CURRICULUM");       // 교육명         
                int vIDX_Col2 = pGrid.GetColumnToIndex("EDUCATION_PERIOD");     // 교육기간
                int vIDX_Col3 = pGrid.GetColumnToIndex("EDU_ORG");              // 시행처
                //int vIDX_Col4 = pGrid.GetColumnToIndex("LICENSE_DATE");         // 취득일자

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
            //상벌사항 
            mPrinting.XLActiveSheet("SourceTab1");
            int vROW = 38;  //인쇄 위치 

            try
            {
                int vIDX_Col1 = pGrid.GetColumnToIndex("RP_DATE");          // 상벌일자
                int vIDX_Col2 = pGrid.GetColumnToIndex("RP_TYPE_NAME");     // 상벌구분
                int vIDX_Col3 = pGrid.GetColumnToIndex("RP_NAME");          // 종류
                int vIDX_Col4 = pGrid.GetColumnToIndex("RP_DESCRIPTION");   // 사유

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

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = mXLColumnAreaSTART; //복사되어질 쉬트의 폭, 시작열
            int vPrintHeaderColumnEND = mXLColumnAreaEND;     //복사되어질 쉬트의 폭, 종료열

            int vCopySumPrintingLine = 0;
            vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mMaxIncrementCopy;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mMaxIncrementCopy, vPrintHeaderColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
                mPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vCopySumPrintingLine;
            //mPrinting.XLPrintPreview();
        }

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste2(int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = mSTART_COL; //복사되어질 쉬트의 폭, 시작열
            int vPrintHeaderColumnEND = mEND_COL;     //복사되어질 쉬트의 폭, 종료열

            int vCopySumPrintingLine = 0;
            vCopySumPrintingLine = pCopySumPrintingLine;

            try
            {
                int vCopyPrintingRowSTART = vCopySumPrintingLine;
                vCopySumPrintingLine = vCopySumPrintingLine + mEND_ROW;
                int vCopyPrintingRowEnd = vCopySumPrintingLine;

                mPrinting.XLActiveSheet("SourceTab1"); //mPrinting.XLActiveSheet(2);
                object vRangeSource = mPrinting.XLGetRange(mSTART_ROW, mSTART_COL,mEND_ROW, mEND_COL); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호

                mPrinting.XLActiveSheet("Destination"); //mPrinting.XLActiveSheet(1);
                object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
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