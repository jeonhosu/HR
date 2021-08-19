using System;
using System.Collections.Generic;
using System.Text;

namespace HRMF0705
{
    class XLPrinting_3
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        //private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        //private int mPrintingLineSTART = 8;  //Line

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
        private int mIncrementCopyMAX = 74;  // 1page : 61, 2page : 122, 3page : 183 - 복사되어질 행의 범위

        private int mCopyColumnSTART = 1;    //복사되어  진 행 누적 수
        private int mCopyColumnEND = 41;     //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

        private string mSend_ORG = string.Empty;
        private string mPrint_COUNT = string.Empty;

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

        public XLPrinting_3(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        #region ----- Array Set -----

        private void SetArray(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_IN_EARNER_DED_TAX, out int[] pGDColumn)
        {
            pGDColumn = new int[303];

            pGDColumn[0] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("YEAR_YYYY");              // 귀속년도
            pGDColumn[1] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("CORP_NAME");              // 법인명(상호)   
            pGDColumn[2] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("VAT_NUMBER");             // 사업자번호     
            pGDColumn[3] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("CORP_ADDRESS");           // 사업자주소               
            pGDColumn[4] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("NAME");                   // 성명           
            pGDColumn[5] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("REPRE_NUM");              // 주민번호 
            pGDColumn[6] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("JOIN_DATE");               // 입사일
            pGDColumn[7] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("RETIRE_DATE");             // 퇴사일

            pGDColumn[8] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("NATIONALITY_TYPE_DESC");   // 내외국인구분
            pGDColumn[9] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("NATION_NAME");             // 국가
            pGDColumn[10] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ISO_NATION_CODE");         // 국가코드
            pGDColumn[11] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("DED_FAMILY_COUNT");        // 공제대상
            pGDColumn[12] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("DED_CHILD_COUNT");         // 다자녀
            pGDColumn[13] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("DECREASE_YN");             // 감면여부
            pGDColumn[14] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("DECREASE_REASON");         // 감면규정
            pGDColumn[15] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("DECREASE_PERIOD");         // 감면기간

            //--------------------------------------------------------------------------------------------------------------
            pGDColumn[16] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_01");           // 지급연월
            pGDColumn[17] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_01");           // 급여액
            pGDColumn[18] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_01");         // 상여액         
            pGDColumn[19] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_01");           // 급여계
            pGDColumn[20] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_01");     // 간이세액표적용구간
            pGDColumn[21] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_01");      // 소득세
            pGDColumn[22] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_01");       // 그외소득세         
            pGDColumn[23] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_01");    // 소득세액     
            pGDColumn[24] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_01");     // 지방소득세

            pGDColumn[25] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_02");           // 지급연월
            pGDColumn[26] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_02");           // 급여액
            pGDColumn[27] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_02");         // 상여액         
            pGDColumn[28] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_02");           // 급여계
            pGDColumn[29] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_02");     // 간이세액표적용구간
            pGDColumn[30] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_02");      // 소득세
            pGDColumn[31] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_02");       // 그외소득세         
            pGDColumn[32] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_02");    // 소득세액     
            pGDColumn[33] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_02");     // 지방소득세

            pGDColumn[34] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_03");           // 지급연월
            pGDColumn[35] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_03");           // 급여액
            pGDColumn[36] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_03");         // 상여액         
            pGDColumn[37] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_03");           // 급여계
            pGDColumn[38] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_03");     // 간이세액표적용구간
            pGDColumn[39] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_03");      // 소득세
            pGDColumn[40] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_03");       // 그외소득세         
            pGDColumn[41] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_03");    // 소득세액     
            pGDColumn[42] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_03");     // 지방소득세

            pGDColumn[43] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_04");           // 지급연월
            pGDColumn[44] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_04");           // 급여액
            pGDColumn[45] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_04");         // 상여액         
            pGDColumn[46] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_04");           // 급여계
            pGDColumn[47] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_04");     // 간이세액표적용구간
            pGDColumn[48] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_04");      // 소득세
            pGDColumn[49] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_04");       // 그외소득세         
            pGDColumn[50] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_04");    // 소득세액     
            pGDColumn[51] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_04");     // 지방소득세

            pGDColumn[52] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_05");           // 지급연월
            pGDColumn[53] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_05");           // 급여액
            pGDColumn[54] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_05");         // 상여액         
            pGDColumn[55] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_05");           // 급여계
            pGDColumn[56] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_05");     // 간이세액표적용구간
            pGDColumn[57] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_05");      // 소득세
            pGDColumn[58] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_05");       // 그외소득세         
            pGDColumn[59] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_05");    // 소득세액     
            pGDColumn[60] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_05");     // 지방소득세

            pGDColumn[61] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_06");           // 지급연월
            pGDColumn[62] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_06");           // 급여액
            pGDColumn[63] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_06");         // 상여액         
            pGDColumn[64] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_06");           // 급여계
            pGDColumn[65] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_06");     // 간이세액표적용구간
            pGDColumn[66] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_06");      // 소득세
            pGDColumn[67] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_06");       // 그외소득세         
            pGDColumn[68] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_06");    // 소득세액     
            pGDColumn[69] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_06");     // 지방소득세

            pGDColumn[70] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_07");           // 지급연월
            pGDColumn[71] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_07");           // 급여액
            pGDColumn[72] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_07");         // 상여액         
            pGDColumn[73] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_07");           // 급여계
            pGDColumn[74] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_07");     // 간이세액표적용구간
            pGDColumn[75] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_07");      // 소득세
            pGDColumn[76] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_07");       // 그외소득세         
            pGDColumn[77] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_07");    // 소득세액     
            pGDColumn[78] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_07");     // 지방소득세

            pGDColumn[79] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_08");           // 지급연월
            pGDColumn[80] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_08");           // 급여액
            pGDColumn[81] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_08");         // 상여액         
            pGDColumn[82] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_08");           // 급여계
            pGDColumn[83] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_08");     // 간이세액표적용구간
            pGDColumn[84] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_08");      // 소득세
            pGDColumn[85] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_08");       // 그외소득세         
            pGDColumn[86] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_08");    // 소득세액     
            pGDColumn[87] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_08");     // 지방소득세

            pGDColumn[88] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_09");           // 지급연월
            pGDColumn[89] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_09");           // 급여액
            pGDColumn[90] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_09");         // 상여액         
            pGDColumn[91] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_09");           // 급여계
            pGDColumn[92] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_09");     // 간이세액표적용구간
            pGDColumn[93] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_09");      // 소득세
            pGDColumn[94] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_09");       // 그외소득세         
            pGDColumn[95] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_09");    // 소득세액     
            pGDColumn[96] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_09");     // 지방소득세

            pGDColumn[97] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_10");           // 지급연월
            pGDColumn[98] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_10");           // 급여액
            pGDColumn[99] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_10");         // 상여액         
            pGDColumn[100] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_10");          // 급여계
            pGDColumn[101] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_10");    // 간이세액표적용구간
            pGDColumn[102] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_10");     // 소득세
            pGDColumn[103] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_10");      // 그외소득세         
            pGDColumn[104] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_10");   // 소득세액     
            pGDColumn[105] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_10");    // 지방소득세

            pGDColumn[106] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_11");           // 지급연월
            pGDColumn[107] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_11");           // 급여액
            pGDColumn[108] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_11");         // 상여액         
            pGDColumn[109] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_11");          // 급여계
            pGDColumn[110] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_11");    // 간이세액표적용구간
            pGDColumn[111] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_11");     // 소득세
            pGDColumn[112] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_11");      // 그외소득세         
            pGDColumn[113] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_11");   // 소득세액     
            pGDColumn[114] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_11");    // 지방소득세

            pGDColumn[115] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_YYYYMM_12");           // 지급연월
            pGDColumn[116] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PAY_AMOUNT_12");           // 급여액
            pGDColumn[117] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("BONUS_AMOUNT_12");         // 상여액         
            pGDColumn[118] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUM_AMOUNT_12");          // 급여계
            pGDColumn[119] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_SECTION_12");    // 간이세액표적용구간
            pGDColumn[120] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TEMP_TAX_AMOUNT_12");     // 소득세
            pGDColumn[121] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ETC_TAX_AMOUNT_12");      // 그외소득세         
            pGDColumn[122] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT_12");   // 소득세액     
            pGDColumn[123] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("LOCAL_TAX_AMOUNT_12");    // 지방소득세

            pGDColumn[124] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_PAY_AMOUNT");          // 지급연월
            pGDColumn[125] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_BONUS_AMOUNT");        // 급여액
            pGDColumn[126] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_SUM_AMOUNT");          // 상여액         
            pGDColumn[127] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TEMP_TAX_SECTION");    // 급여계
            pGDColumn[128] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TEMP_TAX_AMOUNT");     // 간이세액표적용구간
            pGDColumn[129] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_ETC_TAX_AMOUNT");      // 소득세
            pGDColumn[130] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT");   // 그외소득세         
            pGDColumn[131] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_LOCAL_TAX_AMOUNT");    // 소득세액     
            //-------------------------------------------------------------------------------------------------------------- 
            pGDColumn[132] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_01");       // 국외근로
            pGDColumn[133] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_01");            // 연장
            pGDColumn[134] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_01");          // 보육         
            pGDColumn[135] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_01");          // 소계
            pGDColumn[136] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_01");         // 합계
            pGDColumn[137] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_01");           // 차량유지비
            pGDColumn[138] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_01");           // 기타         
            pGDColumn[139] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_01");         // 합계

            pGDColumn[140] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_02");       // 국외근로
            pGDColumn[141] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_02");            // 연장
            pGDColumn[142] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_02");          // 보육         
            pGDColumn[143] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_02");          // 소계
            pGDColumn[144] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_02");         // 합계
            pGDColumn[145] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_02");           // 차량유지비
            pGDColumn[146] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_02");           // 기타         
            pGDColumn[147] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_02");         // 합계

            pGDColumn[148] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_03");       // 국외근로
            pGDColumn[149] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_03");            // 연장
            pGDColumn[150] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_03");          // 보육         
            pGDColumn[151] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_03");          // 소계
            pGDColumn[152] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_03");         // 합계
            pGDColumn[153] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_03");           // 차량유지비
            pGDColumn[154] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_03");           // 기타         
            pGDColumn[155] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_03");         // 합계

            pGDColumn[156] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_04");       // 국외근로
            pGDColumn[157] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_04");            // 연장
            pGDColumn[158] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_04");          // 보육         
            pGDColumn[159] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_04");          // 소계
            pGDColumn[160] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_04");         // 합계
            pGDColumn[161] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_04");           // 차량유지비
            pGDColumn[162] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_04");           // 기타         
            pGDColumn[163] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_04");         // 합계

            pGDColumn[164] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_05");       // 국외근로
            pGDColumn[165] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_05");            // 연장
            pGDColumn[166] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_05");          // 보육         
            pGDColumn[167] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_05");          // 소계
            pGDColumn[168] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_05");         // 합계
            pGDColumn[169] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_05");           // 차량유지비
            pGDColumn[170] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_05");           // 기타         
            pGDColumn[171] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_05");         // 합계

            pGDColumn[172] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_06");       // 국외근로
            pGDColumn[173] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_06");            // 연장
            pGDColumn[174] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_06");          // 보육         
            pGDColumn[175] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_06");          // 소계
            pGDColumn[176] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_06");         // 합계
            pGDColumn[177] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_06");           // 차량유지비
            pGDColumn[178] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_06");           // 기타         
            pGDColumn[179] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_06");         // 합계

            pGDColumn[180] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_07");       // 국외근로
            pGDColumn[181] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_07");            // 연장
            pGDColumn[182] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_07");          // 보육         
            pGDColumn[183] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_07");          // 소계
            pGDColumn[184] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_07");         // 합계
            pGDColumn[185] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_07");           // 차량유지비
            pGDColumn[186] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_07");           // 기타         
            pGDColumn[187] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_07");         // 합계

            pGDColumn[188] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_08");       // 국외근로
            pGDColumn[189] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_08");            // 연장
            pGDColumn[190] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_08");          // 보육         
            pGDColumn[191] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_08");          // 소계
            pGDColumn[192] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_08");         // 합계
            pGDColumn[193] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_08");           // 차량유지비
            pGDColumn[194] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_08");           // 기타         
            pGDColumn[195] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_08");         // 합계

            pGDColumn[196] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_09");       // 국외근로
            pGDColumn[197] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_09");            // 연장
            pGDColumn[198] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_09");          // 보육         
            pGDColumn[199] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_09");          // 소계
            pGDColumn[200] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_09");         // 합계
            pGDColumn[201] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_09");           // 차량유지비
            pGDColumn[202] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_09");           // 기타         
            pGDColumn[203] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_09");         // 합계

            pGDColumn[204] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_10");       // 국외근로
            pGDColumn[205] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_10");            // 연장
            pGDColumn[206] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_10");          // 보육         
            pGDColumn[207] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_10");          // 소계
            pGDColumn[208] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_10");         // 합계
            pGDColumn[209] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_10");           // 차량유지비
            pGDColumn[210] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_10");           // 기타         
            pGDColumn[211] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_10");         // 합계

            pGDColumn[212] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_11");       // 국외근로
            pGDColumn[213] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_11");            // 연장
            pGDColumn[214] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_11");          // 보육         
            pGDColumn[215] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_11");          // 소계
            pGDColumn[216] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_11");         // 합계
            pGDColumn[217] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_11");           // 차량유지비
            pGDColumn[218] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_11");           // 기타         
            pGDColumn[219] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_11");         // 합계

            pGDColumn[220] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OUTSIDE_12");       // 국외근로
            pGDColumn[221] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_OT_12");            // 연장
            pGDColumn[222] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_BABY_12");          // 보육         
            pGDColumn[223] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_PART_12");          // 소계
            pGDColumn[224] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_1_12");         // 합계
            pGDColumn[225] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_CAR_12");           // 차량유지비
            pGDColumn[226] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_ETC_12");           // 기타         
            pGDColumn[227] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TAX_FREE_SUM_2_12");         // 합계

            pGDColumn[228] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_OUTSIDE");       // 국외근로
            pGDColumn[229] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_OT");            // 연장
            pGDColumn[230] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_BABY");          // 보육         
            pGDColumn[231] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_PART");          // 소계
            pGDColumn[232] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_SUM_1");         // 합계
            pGDColumn[233] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_CAR");           // 차량유지비
            pGDColumn[234] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_ETC");           // 기타         
            pGDColumn[235] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_TAX_FREE_SUM_2");         // 합계

            //--------------------------------------------------------------------------------------------------------------
            pGDColumn[236] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_01");     // 소득세         
            pGDColumn[237] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_01");  // 주민세         
            pGDColumn[238] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_01");      // 연금보험       
            pGDColumn[239] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_01");       // 건강보험       
            pGDColumn[240] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_01");          // 고용보험       

            pGDColumn[241] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_02");     // 소득세         
            pGDColumn[242] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_02");  // 주민세         
            pGDColumn[243] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_02");      // 연금보험       
            pGDColumn[244] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_02");       // 건강보험       
            pGDColumn[245] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_02");          // 고용보험       

            pGDColumn[246] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_03");     // 소득세         
            pGDColumn[247] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_03");  // 주민세          
            pGDColumn[248] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_03");      // 연금보험       
            pGDColumn[249] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_03");       // 건강보험       
            pGDColumn[250] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_03");          // 고용보험       

            pGDColumn[251] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_04");     // 소득세         
            pGDColumn[252] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_04");  // 주민세         
            pGDColumn[253] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_04");      // 연금보험       
            pGDColumn[254] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_04");       // 건강보험       
            pGDColumn[255] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_04");          // 고용보험       

            pGDColumn[256] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_05");     // 소득세         
            pGDColumn[257] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_05");  // 주민세         
            pGDColumn[258] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_05");      // 연금보험       
            pGDColumn[259] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_05");       // 건강보험       
            pGDColumn[260] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_05");          // 고용보험       

            pGDColumn[261] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_06");     // 소득세         
            pGDColumn[262] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_06");  // 주민세         
            pGDColumn[263] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_06");      // 연금보험       
            pGDColumn[264] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_06");       // 건강보험       
            pGDColumn[265] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_06");          // 고용보험       

            pGDColumn[266] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_07");     // 소득세         
            pGDColumn[267] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_07");  // 주민세         
            pGDColumn[268] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_07");      // 연금보험       
            pGDColumn[269] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_07");       // 건강보험       
            pGDColumn[270] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_07");          // 고용보험       

            pGDColumn[271] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_08");     // 소득세         
            pGDColumn[272] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_08");  // 주민세         
            pGDColumn[273] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_08");      // 연금보험       
            pGDColumn[274] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_08");       // 건강보험       
            pGDColumn[275] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_08");          // 고용보험       

            pGDColumn[276] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_09");     // 소득세         
            pGDColumn[277] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_09");  // 주민세         
            pGDColumn[278] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_09");      // 연금보험       
            pGDColumn[279] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_09");       // 건강보험       
            pGDColumn[280] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_09");          // 고용보험       

            pGDColumn[281] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_10");     // 소득세         
            pGDColumn[282] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_10");  // 주민세         
            pGDColumn[283] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_10");      // 연금보험       
            pGDColumn[284] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_10");       // 건강보험       
            pGDColumn[285] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_10");          // 고용보험       

            pGDColumn[286] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_11");     // 소득세         
            pGDColumn[287] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_11");  // 주민세         
            pGDColumn[288] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_11");      // 연금보험       
            pGDColumn[289] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_11");       // 건강보험       
            pGDColumn[290] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_11");          // 고용보험       

            pGDColumn[291] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT_12");     // 소득세         
            pGDColumn[292] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT_12");  // 주민세         
            pGDColumn[293] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("ANNUITY_IN_AMT_12");      // 연금보험       
            pGDColumn[294] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("HEALTH_IN_AMT_12");       // 건강보험       
            pGDColumn[295] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("UMP_IN_AMT_12");          // 고용보험

            pGDColumn[296] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_SUBT_IN_TAX_AMT");   // 소득세 계         
            pGDColumn[297] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_SUBT_LOCAL_TAX_AMT");// 주민세 계         
            pGDColumn[298] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_ANNUITY_IN_AMT");    // 연금보험 계       
            pGDColumn[299] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_HEALTH_IN_AMT");     // 건강보험 계       
            pGDColumn[300] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("TOTAL_UMP_IN_AMT");        // 고용보험 계 

            pGDColumn[301] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("PRINT_DATE");              // 출력날짜 
            pGDColumn[302] = pGrid_IN_EARNER_DED_TAX.GetColumnToIndex("AGENT_NAME");              // 징수의무자 
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
                mAppInterface.OnAppMessageEvent(ex.Message);
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


        #endregion;

        #region ----- Line Write Method -----

        #region -----XLLINE -----

        private int XLLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGridPRINT_INCOME_TAX, int pGridRow, int[] pGDColumn)
        {
            int vXLine = 0; // 엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            string vPerson_Info = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Sheet1");

                //-------------------------------------------------------------------
                vXLine = 3;
                //-------------------------------------------------------------------
                //귀속연도.
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[0]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = 5;
                //-------------------------------------------------------------------
                // 법인명(상호)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[1]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;  //6
                //-------------------------------------------------------------------
                // 사업자등록번호
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[2]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //7
                //-------------------------------------------------------------------
                // 근무처
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[3]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //8
                //-------------------------------------------------------------------
                // 성명
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[4]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    vPerson_Info = vConvertString;
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 주민등록번호
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[5]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    vPerson_Info = string.Format("{0}({1})", vPerson_Info, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 23;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //페이지 상단 사원정보 인쇄//
                mPrinting.XLSetCell(1, 23, vPerson_Info);
                mPrinting.XLSetCell(33, 23, vPerson_Info);
                mPrinting.XLSetCell(56, 23, vPerson_Info);

                // 입사일자
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[6]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //9
                //-------------------------------------------------------------------
                // 입사일자
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[7]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //10
                //-------------------------------------------------------------------
                // 내외국인구분
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[8]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국가
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[9]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 23;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국가코드
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[10]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //11
                //-------------------------------------------------------------------
                // 공제대상
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[11]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //12
                //-------------------------------------------------------------------
                // 다자녀
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[12]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //13
                //-------------------------------------------------------------------
                // 감면여부
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[13]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 감면규정
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[14]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 감면기간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[15]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //---------------------------- 급상여 내역 --------------------------
                //-------------------------------------------------------------------
                vXLine = 18;
                //-------------------------------------------------------------------
                // 지급연월(1월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[16]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[17]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[18]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[19]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[20]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[21]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[22]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[23]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[24]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(2월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[25]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[26]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[27]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[28]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[29]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[30]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[31]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[32]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[33]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(3월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[34]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[35]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[36]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[37]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[38]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[39]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[40]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[41]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[42]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(4월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[43]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[44]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[45]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[46]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[47]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[48]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[49]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[50]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[51]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(5월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[52]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[53]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[54]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[55]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[56]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[57]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[58]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[59]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[60]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(6월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[61]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[62]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[63]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[64]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[65]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[66]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[67]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[68]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[69]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(7월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[70]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[71]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[72]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[73]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[74]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[75]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[76]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[77]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[78]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(8월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[79]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[80]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[81]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[82]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[83]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[84]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[85]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[86]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[87]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(9월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[88]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[89]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[90]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[91]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[92]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[93]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[94]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[95]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[96]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(10월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[97]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[98]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[99]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[100]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[101]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[102]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[103]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[104]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[105]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(11월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[106]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[107]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[108]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[109]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[110]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[111]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[112]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[113]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[114]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(12월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[115]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[116]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[117]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[118]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[119]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[120]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[121]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[122]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[123]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 급여액(합계)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[124]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 상여액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[125]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 급여계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[126]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 간이세액표적용구간
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[127]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[128]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 그외소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[129]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 소득세액
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[130]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 36;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[131]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //---------------------------- 비과세 소득 --------------------------
                //-------------------------------------------------------------------
                vXLine = 40;
                //-------------------------------------------------------------------
                // 지급연월(1월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[16]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[132]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[133]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[134]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[135]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[136]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[137]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[139]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(2월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[25]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[140]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[141]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[142]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[143]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[144]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[145]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[147]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(3월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[34]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[148]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[149]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[150]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[151]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[152]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[153]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[155]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(4월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[43]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[156]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[157]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[158]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[159]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[160]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[161]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[163]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(5월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[52]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[164]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[165]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[166]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[167]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[168]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[169]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[171]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(6월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[61]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[172]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[173]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[174]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[175]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[176]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[177]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[179]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(7월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[70]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[180]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[181]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[182]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[183]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[184]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[185]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[187]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(8월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[79]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[188]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[189]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[190]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[191]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[192]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[193]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[195]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(9월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[88]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[196]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[197]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[198]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[199]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[200]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[201]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[203]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(10월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[97]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[204]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[205]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[206]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[207]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[208]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[209]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[211]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(11월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[106]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[212]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[213]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[214]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[215]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[216]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[217]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[219]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 지급연월(12월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[115]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 국외근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[220]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[221]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[222]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[223]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[224]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[225]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[227]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 국외근로(합계)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[228]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 야간근로
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[229]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 보육
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[230]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //// 소계
                //vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[231]);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumn = 22;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 비과세합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[232]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 자가운전비
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[233]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[235]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //------------------------ 근로소득원천징수액 -------------------------//
                //-------------------------------------------------------------------
                vXLine = 60;
                //-------------------------------------------------------------------
                // 소득세(1월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[236]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[237]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[238]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[239]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[240]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(2월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[241]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[242]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[243]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[244]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[245]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(3월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[246]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[247]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[248]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[249]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[250]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(4월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[251]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[252]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[253]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[254]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[255]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(5월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[256]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[257]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[258]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[259]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[260]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(6월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[261]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[262]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[263]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[264]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[265]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(7월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[266]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[267]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[268]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[269]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[270]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(8월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[271]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[272]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[273]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[274]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[275]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(9월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[276]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[277]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[278]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[279]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[280]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(10월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[281]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[282]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[283]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[284]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[285]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(11월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[286]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[287]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[288]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[289]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[290]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세(12월)
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[291]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[292]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[293]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[294]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[295]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 소득세 합계
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[296]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 지방소득세
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[297]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 연금보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[298]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 건강보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[299]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 고용보험
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[300]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 33;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 출력 날짜
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[301]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                // 징수의무자
                vObject = pGridPRINT_INCOME_TAX.GetCellValue(pGridRow, pGDColumn[302]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
            }

            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            return vXLine; ; ;
        }

        #endregion;

        #endregion;

        #region ----- Excel Main Wirte  Method Backup----

        public int WriteMain(InfoSummit.Win.ControlAdv.ISGridAdvEx pGridIN_EARNER_DED_TAX)
        {
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;

            int vTotalRow = pGridIN_EARNER_DED_TAX.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            SetArray(pGridIN_EARNER_DED_TAX, out vGDColumn);
            mPrinting.XLActiveSheet("Source1");

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                pGridIN_EARNER_DED_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "Sheet1");

                pGridIN_EARNER_DED_TAX.CurrentCellMoveTo(vRow, 0);
                pGridIN_EARNER_DED_TAX.Focus();
                pGridIN_EARNER_DED_TAX.CurrentCellActivate(vRow, 0);

                vPrintingLine = XLLine(pGridIN_EARNER_DED_TAX, vRow, vGDColumn);
            }

            return mPageNumber;

            //---------------------------------------------------------------------------------------------------
            // 설  명 : Form에서 성명을 선택하지 않았을 시, '전 직원' 출력이 가능하도록 구현한 소스 코드입니다.
            //          향후, 1인 출력이 아닌 전체 출력으로 변경해야 할 경우 사용하세요.
            // 날  짜 : 2011. 6. 14(화)
            // 작성자 : 이선희J
            //---------------------------------------------------------------------------------------------------
            /*
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = gridIN_EARNER_DED_TAX.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 30;
            int vCountPrinting = 0;

            SetArray(gridIN_EARNER_DED_TAX, out vGDColumn, out vXLColumn);
            mPrinting.XLActiveSheet("SourceTab1");

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                gridIN_EARNER_DED_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                    vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                    gridIN_EARNER_DED_TAX.CurrentCellMoveTo(vRow, 0);
                    gridIN_EARNER_DED_TAX.Focus();
                    gridIN_EARNER_DED_TAX.CurrentCellActivate(vRow, 0);

                    vPrintingLine = XLLine(gridIN_EARNER_DED_TAX, vRow, vPrintingLine, vGDColumn, vXLColumn, "SRC_TAB1");

                    if (vSecondPrinting < vCountPrinting)
                    {
                        Printing(1, vSecondPrinting);

                        mPrinting.XLOpenFileClose();
                        isOpen = XLFileOpen();

                        vCountPrinting = 0;
                        vPrintingLine = 1;
                        mCopyLineSUM = 1;
                    }
                    else if (vTotalRow == vRowCount)
                    {
                        Printing(1, vSecondPrinting);
                    }
                }
            }
            mPrinting.XLOpenFileClose();            

            return mPageNumber;
            */
        }

        #endregion;

        #region ----- Header Write Method ----

        private void XLHeader(string pDate)
        {
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                vXLine = 53;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, pDate);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }


            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                vXLine = 61;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, pDate);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //첫번째 페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, string pCourse)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;

            pPrinting.XLActiveSheet("Source1");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet("Sheet1");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber = 3; //페이지 번호

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPrinting(pPageSTART, pPageEND);
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder.ToString(), vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- PDF Method ----

        public void PDF(string pSaveFileName)
        {
            try
            {
                bool isSuccess = mPrinting.XLSaveAs_PDF(pSaveFileName);  // DELETED, BY MJSHIN
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Delete Sheet Method ----

        public void DeleteSheet()
        {
            bool isSuccess = false;

            try
            {
                isSuccess = mPrinting.XLDeleteSheet("Source1");

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
